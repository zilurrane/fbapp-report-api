package main

import (
	"bytes"
	"context"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"
	"time"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/gin-gonic/gin"
	cors "github.com/rs/cors/wrapper/gin"

	"go.mongodb.org/mongo-driver/bson"
	"go.mongodb.org/mongo-driver/bson/primitive"
	"go.mongodb.org/mongo-driver/mongo"
	"go.mongodb.org/mongo-driver/mongo/options"
	"go.mongodb.org/mongo-driver/mongo/readpref"
)

// Department Department
type Department struct {
	Code string
	Name string
}

// Class Class
type Class struct {
	Code string
	Name string
}

// Meta Meta
type Meta struct {
	Class      Class
	Department Department
}

// FacultyFeedback FacultyFeedback
type FacultyFeedback struct {
	Name     string
	Feedback int8
}

// FacultyComparionReportRequestBody FacultyComparionReportRequestBody
type FacultyComparionReportRequestBody struct {
	Meta Meta
	Data []FacultyFeedback
}

// Feedback
type Feedback struct {
	FbNo           int       `json:"fbNo"`
	DepartmentCode string    `json:"departmentCode"`
	ClassCode      string    `json:"classCode"`
	Faculty        string    `json:"faculty"`
	CreatedDate    time.Time `json:"createdDate"`
	Feedback       bson.M    `bson:"feedback"`
}

type FeedbackParameter struct {
	ID   primitive.ObjectID `bson:"_id, omitempty"`
	Code string             `json:"code"`
}

func main() {
	r := gin.Default()

	handler := cors.New(cors.Options{
		AllowedOrigins:   []string{"*"},
		AllowCredentials: true,
		AllowedHeaders:   []string{"*"},
		ExposedHeaders:   []string{"Content-Disposition"},
	})

	r.Use(handler)
	r.GET("/ping", healthStatusPing)
	r.POST("/report/faculty/feedback/comparison", getFacultyFeedbackComparisonReport)
	r.GET("/report/feedback", getFeedback)
	var port string
	if port = os.Getenv("PORT"); len(port) == 0 {
		port = "8080"
	}
	r.Run(":" + port)
}

func healthStatusPing(c *gin.Context) {
	c.JSON(200, gin.H{
		"message": "pong",
	})
}

func toChar(i int) string {
	return string('A' - 1 + i)
}

func getFeedback(c *gin.Context) {

	tenantId := c.Request.Header.Get("TenantId")

	clientOptions := options.Client().ApplyURI("mongodb+srv://adminbhau:adminbhau@cluster0.nokmu.mongodb.net")
	client, err := mongo.NewClient(clientOptions)

	//Set up a context required by mongo.Connect
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	err = client.Connect(ctx)

	//To close the connection at the end
	defer cancel()

	err = client.Ping(context.Background(), readpref.Primary())
	if err != nil {
		log.Fatal("Couldn't connect to the database", err)
	} else {
		log.Println("Connected!")
	}
	db := client.Database("fbapp")

	cursor, err := db.Collection("feedbacks").Find(context.TODO(), bson.M{"tenantId": bson.M{"$eq": tenantId}})

	if err != nil {
		log.Printf("Error while getting all feedbacks, Reason: %v\n", err)
		c.JSON(http.StatusInternalServerError, gin.H{
			"status":  http.StatusInternalServerError,
			"message": "Something went wrong",
		})
		return
	}

	findOptions := options.Find()

	findOptions.SetSort(bson.D{{"createdDate", -1}})

	parameterCursor, err := db.Collection("feedbackparameters").Find(context.TODO(), bson.M{"tenantId": bson.M{"$eq": tenantId}})

	if err != nil {
		log.Printf("Error while getting all feedback parameters, Reason: %v\n", err)
		c.JSON(http.StatusInternalServerError, gin.H{
			"status":  http.StatusInternalServerError,
			"message": "Something went wrong",
		})
		return
	}

	var feedbackParameters []FeedbackParameter

	for parameterCursor.Next(context.TODO()) {
		var parameter FeedbackParameter
		parameterCursor.Decode(&parameter)
		feedbackParameters = append(feedbackParameters, parameter)
	}

	facultyCursor, err := db.Collection("faculties").Find(context.TODO(), bson.M{"tenantId": bson.M{"$eq": tenantId}})

	if err != nil {
		log.Printf("Error while getting all faculties, Reason: %v\n", err)
		c.JSON(http.StatusInternalServerError, gin.H{
			"status":  http.StatusInternalServerError,
			"message": "Something went wrong",
		})
		return
	}

	var faculties = make(map[string]string)

	for facultyCursor.Next(context.TODO()) {
		var faculty bson.M
		facultyCursor.Decode(&faculty)
		faculties[faculty["_id"].(primitive.ObjectID).Hex()] = fmt.Sprintf("%v", faculty["name"])
	}

	sheetName := "Feedback"
	f := excelize.NewFile()

	f.SetSheetName("Sheet1", sheetName)

	headerTitleStyle, headerTitleStyleErr := f.NewStyle(`{"font":{"bold": true}}`)

	if headerTitleStyleErr != nil {
		fmt.Println(headerTitleStyleErr)
	}

	f.SetCellStyle(sheetName, "A1", "K1", headerTitleStyle)

	f.SetCellValue(sheetName, fmt.Sprintf("A%d", 1), "Sr. No.")
	f.SetCellValue(sheetName, fmt.Sprintf("B%d", 1), "Faculty")
	f.SetCellValue(sheetName, fmt.Sprintf("C%d", 1), "Department")
	f.SetCellValue(sheetName, fmt.Sprintf("D%d", 1), "Class")
	f.SetCellValue(sheetName, fmt.Sprintf("E%d", 1), "CreatedAt")

	for index, parameter := range feedbackParameters {
		f.SetCellValue(sheetName, fmt.Sprintf("%s%d", toChar(index+6), 1), parameter.Code)
	}

	var feedbackIndex int = 1

	for cursor.Next(context.TODO()) {
		var feedback Feedback
		cursor.Decode(&feedback)
		feedbackIndex += 1
		f.SetCellValue(sheetName, fmt.Sprintf("A%d", feedbackIndex), feedbackIndex-1)
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", feedbackIndex), faculties[feedback.Faculty])
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", feedbackIndex), feedback.DepartmentCode)
		f.SetCellValue(sheetName, fmt.Sprintf("D%d", feedbackIndex), feedback.ClassCode)
		f.SetCellValue(sheetName, fmt.Sprintf("E%d", feedbackIndex), feedback.CreatedDate)

		for index, parameter := range feedbackParameters {
			f.SetCellValue(sheetName, fmt.Sprintf("%s%d", toChar(index+6), feedbackIndex), feedback.Feedback[parameter.ID.Hex()])
		}
	}

	var b bytes.Buffer
	if err := f.Write(&b); err != nil {
		c.JSON(http.StatusInternalServerError, err.Error())
		return
	}

	downloadName := "Feedback" + ".xlsx"
	c.Header("Content-Description", "File Transfer")
	c.Header("Content-Disposition", "attachment; filename="+downloadName)
	c.Data(http.StatusOK, "application/octet-stream", b.Bytes())
}

func getFacultyFeedbackComparisonReport(c *gin.Context) {

	facultyComparionReportRequestBodyJSON, _ := ioutil.ReadAll(c.Request.Body)

	// []byte(`{"meta":{"department":{"code":"COMP","name":"Computer Engineering"},"class":{"code":"SY","name":"Second Year"}},"data":[{"name":"Prof. V. S. More","feedback":88},{"name":"Prof. Tinku Sharma","feedback":72},{"name":"Mrs. Ileana Mukherjee","feedback":90},{"name":"Prof. Monu Mingle","feedback":88},{"name":"Ms. Hana More","feedback":30}]}`)

	var facultyComparionReportRequest FacultyComparionReportRequestBody
	err := json.Unmarshal(facultyComparionReportRequestBodyJSON, &facultyComparionReportRequest)
	if err != nil {
		log.Println(err)
	}

	sheetName := "FacultyFeedbackComparison"
	f := excelize.NewFile()

	f.SetSheetName("Sheet1", sheetName)

	reportTitleStyle, reportTitleStyleErr := f.NewStyle(`{"font":{"bold": true,"size":16}}`)

	if reportTitleStyleErr != nil {
		fmt.Println(reportTitleStyleErr)
	}

	metaInfoStyle, metaInfoStyleErr := f.NewStyle(`{"font":{"bold": true,"size":12}}`)

	if metaInfoStyleErr != nil {
		fmt.Println(metaInfoStyleErr)
	}

	if err := f.SetColWidth(sheetName, "B", "B", 12); err != nil {
		fmt.Println(err)
		return
	}

	f.SetCellStyle(sheetName, "C2", "C2", reportTitleStyle)
	f.SetCellValue(sheetName, "C2", "Faculty Feedback Comparison")

	f.SetCellStyle(sheetName, "B4", "B4", metaInfoStyle)
	f.SetCellValue(sheetName, "B4", "Department")
	f.SetCellStyle(sheetName, "F4", "F4", metaInfoStyle)
	f.SetCellValue(sheetName, "F4", "Class")

	f.SetCellValue(sheetName, "C4", facultyComparionReportRequest.Meta.Department.Name)
	f.SetCellValue(sheetName, "G4", facultyComparionReportRequest.Meta.Class.Name)

	f.SetCellValue(sheetName, "C6", "Feedback")

	var rowNumber int = 6
	for _, value := range facultyComparionReportRequest.Data {
		rowNumber++
		f.SetCellValue(sheetName, fmt.Sprintf("B%d", rowNumber), value.Name)
		f.SetCellValue(sheetName, fmt.Sprintf("C%d", rowNumber), value.Feedback)
	}

	if err := f.AddChart(sheetName, "A5",
		fmt.Sprintf(`
		{
			"type":"col",
			"series":[
				{
					"name":"FacultyFeedbackComparison!$C$6",
					"categories":"FacultyFeedbackComparison!$B$7:$B$%d",
					"values":"FacultyFeedbackComparison!$C$7:$C$%d"
				}
			],
			"format":{
				"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false
			},
			"title":{
				"name":"Feedback"
			},
			"plotarea":{
				"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":true
			},
			"show_blanks_as":"zero"
		}
		`, rowNumber, rowNumber)); err != nil {
		fmt.Println(err)
	}

	var b bytes.Buffer
	if err := f.Write(&b); err != nil {
		c.JSON(http.StatusInternalServerError, err.Error())
		return
	}

	downloadName := "FacultyFeedbackComparison_" + facultyComparionReportRequest.Meta.Department.Code + "_" + facultyComparionReportRequest.Meta.Class.Code + ".xlsx"
	c.Header("Content-Description", "File Transfer")
	c.Header("Content-Disposition", "attachment; filename="+downloadName)
	c.Data(http.StatusOK, "application/octet-stream", b.Bytes())
}
