package main

import (
	"bytes"
	"fmt"
	"net/http"

	"github.com/360EntSecGroup-Skylar/excelize"
	"github.com/gin-gonic/gin"
)

func main() {
	r := gin.Default()
	r.GET("/ping", healthStatusPing)
	r.GET("/report/faculty/feedback/comparison", getFacultyFeedbackComparisonReport)
	r.Run(":9001")
}

func healthStatusPing(c *gin.Context) {
	c.JSON(200, gin.H{
		"message": "pong",
	})
}

func getFacultyFeedbackComparisonReport(c *gin.Context) {
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

	f.SetCellValue(sheetName, "C4", "Computer Engineering")
	f.SetCellValue(sheetName, "G4", "Second Year")

	categories := map[string]string{"B7": "Prof. V. S. More", "B8": "Prof. Mangesh Gosavi", "C6": "Feedback"}
	values := map[string]int{"C7": 63, "C8": 89}

	for k, v := range categories {
		f.SetCellValue(sheetName, k, v)
	}
	for k, v := range values {
		f.SetCellValue(sheetName, k, v)
	}
	if err := f.AddChart(sheetName, "B10",
		`{"type":"col",
	"series":[
		{"name":"FacultyFeedbackComparison!$C$6","categories":"FacultyFeedbackComparison!$B$7:$B$8","values":"FacultyFeedbackComparison!$C$7:$C$8"}],
		"format":{"x_scale":1.0,"y_scale":1.0,"x_offset":15,"y_offset":10,"print_obj":true,"lock_aspect_ratio":false,"locked":false},
		"title":{"name":"Feedback"},
		"plotarea":{"show_bubble_size":true,"show_cat_name":false,"show_leader_lines":false,"show_percent":true,"show_series_name":false,"show_val":true},
		"show_blanks_as":"zero"}
		`); err != nil {
		fmt.Println(err)
	}

	var b bytes.Buffer
	if err := f.Write(&b); err != nil {
		c.JSON(http.StatusInternalServerError, err.Error())
		return
	}

	downloadName := "FacultyFeedbackComparison.xlsx"
	c.Header("Content-Description", "File Transfer")
	c.Header("Content-Disposition", "attachment; filename="+downloadName)
	c.Data(http.StatusOK, "application/octet-stream", b.Bytes())
}
