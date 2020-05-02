package main

import (
	"bytes"
	"encoding/json"
	"fmt"
	"io/ioutil"
	"log"
	"net/http"
	"os"

	"github.com/360EntSecGroup-Skylar/excelize/v2"
	"github.com/gin-gonic/gin"
	cors "github.com/rs/cors/wrapper/gin"
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

func main() {
	r := gin.Default()
	r.Use(cors.Default())
	r.GET("/ping", healthStatusPing)
	r.POST("/report/faculty/feedback/comparison", getFacultyFeedbackComparisonReport)
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
