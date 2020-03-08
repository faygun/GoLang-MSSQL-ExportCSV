package main

import (
	"database/sql"
	"fmt"
	"os"

	_ "github.com/denisenkom/go-mssqldb"
	"github.com/tealeg/xlsx"
)

func main() {
	fmt.Println("process started.")
	fmt.Println("process in progress.")

	db, err := initDB()

	defer db.Close()

	if err != nil {
		panic("connect step issue: " + err.Error())
	}

	result, err := readDb(db)

	defer result.Close()

	if err != nil {
		panic("read step issue: " + err.Error())
	}

	numbers := fillArray(result)

	convertToCvs(numbers)

	fmt.Println("process finished.")

}

func initDB() (*sql.DB, error) {
	var (
		server   = "your_server"
		port     = 1433
		user     = "your_user"
		password = "your_password"
		database = "your_db_name"
	)

	connString := fmt.Sprintf("server=%s;user id=%s;password=%s;port=%d;database=%s;",
		server, user, password, port, database)

	db, err := sql.Open("mssql", connString)
	if err != nil {
		panic(err.Error())
	}

	fmt.Printf("Connected!\n")

	return db, err
}

func readDb(db *sql.DB) (*sql.Rows, error) {
	query := "your_query"
	rows, err := db.Query(query)

	if err != nil {
		panic(err.Error())
	}

	return rows, err
}

func fillArray(rows *sql.Rows) []string {
	numbers := []string{}
	for rows.Next() {
		var number string
		err := rows.Scan(&number)
		if err != nil {
			panic(err.Error())
		}

		numbers = append(numbers, number)
	}

	return numbers
}

func convertToCvs(numbers []string) {
	path := "numbers.xlsx"

	if numbers == nil {
		panic("array is empty")
	}

	_, err := os.Stat(path)

	if os.IsExist(err) {
		err := os.Remove(path)

		if err != nil {
			panic(err.Error())
		}
	}

	var file *xlsx.File
	var sheet *xlsx.Sheet
	var row *xlsx.Row
	var cell *xlsx.Cell

	file = xlsx.NewFile()
	sheet, err = file.AddSheet("Sheet1")
	if err != nil {
		panic(err.Error())
	}

	for _, num := range numbers {
		row = sheet.AddRow()
		cell = row.AddCell()
		cell.Value = string(num)
	}

	err = file.Save(path)

	if err != nil {
		panic(err.Error())
	}

}
