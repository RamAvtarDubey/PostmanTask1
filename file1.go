package main

import (
	"fmt"
	"log"
	"os"
	"sort"
	"strconv"
	"strings"

	"github.com/xuri/excelize/v2"
)

func main() {

	if len(os.Args) < 2 {
		fmt.Println("Please give two arguments, go filename and excel file path")
		return
	}
	filepath := os.Args[1]

	var validatedLists [][]string
	var invalidLists [][]string

	f, err := excelize.OpenFile(filepath)
	if err != nil {
		fmt.Println("Error opening file, Please make sure you have the currect filepaths")
		return
	}

	rows, err := f.GetRows("CSF111_202425_01_GradeBook")

	if err != nil {
		fmt.Println("Error in reading rows of the file. Please make sure the file contains valid data")
	}

	rows = rows[1:]

	for k := range rows {

		var sumOfPreCT float64 = 0.0

		for i := 4; i <= 7; i++ {
			a, err := strconv.ParseFloat(rows[k][i], 64)
			if err != nil {
				log.Fatal(err)
			}
			sumOfPreCT = sumOfPreCT + a
		}

		// fmt.Println("sumOfPreCT:", sumOfPreCT)

		e, err := strconv.ParseFloat(rows[k][8], 64)
		if err != nil {
			log.Fatal(err)
		}

		o, err := strconv.ParseFloat(rows[k][9], 64)
		if err != nil {
			log.Fatal(err)
		}

		finalsum := sumOfPreCT + o

		q, err := strconv.ParseFloat(rows[k][10], 64)
		if err != nil {
			log.Fatal(err)
		}

		if sumOfPreCT == e && finalsum == q {
			validatedLists = append(validatedLists, rows[k])
		} else {
			invalidLists = append(invalidLists, rows[k])

		}

		//

	}
	if invalidLists != nil {
		fmt.Println("\n\n The following dataset is not a valid set of marks obtained and may include errors :")
		for a := range invalidLists {
			fmt.Println(invalidLists[a])
		}
	}

	// Calculating average now

	quizScore := 0.0
	midsemScore := 0.0
	labtestScore := 0.0
	weeklyScore := 0.0
	prectScore := 0.0
	compreScore := 0.0
	totalScore := 0.0

	A3total := 0.0
	A4total := 0.0
	A5total := 0.0
	A7total := 0.0
	A8total := 0.0
	ADtotal := 0.0
	numberOfStudentsA3 := 0
	numberOfStudentsA4 := 0
	numberOfStudentsA5 := 0
	numberOfStudentsA7 := 0
	numberOfStudentsA8 := 0
	numberOfStudentsAD := 0

	for k := range validatedLists {
		valueOfQuizMarks, err := strconv.ParseFloat(validatedLists[k][4], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOfmidsemMarks, err := strconv.ParseFloat(validatedLists[k][5], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOflabtestMarks, err := strconv.ParseFloat(validatedLists[k][6], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOfweeklyMarks, err := strconv.ParseFloat(validatedLists[k][7], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOfPreCTMarks, err := strconv.ParseFloat(validatedLists[k][8], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOfCompmreMarks, err := strconv.ParseFloat(validatedLists[k][9], 64)
		if err != nil {
			log.Fatal(err)
		}
		valueOfTotalMarks, err := strconv.ParseFloat(validatedLists[k][10], 64)
		if err != nil {
			log.Fatal(err)
		}

		if strings.HasPrefix(validatedLists[k][3], "2024A3") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			A3total += floatConvertScore
			numberOfStudentsA3 += 1
		}

		if strings.HasPrefix(validatedLists[k][3], "2024A4") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			A4total += float64(floatConvertScore)
			numberOfStudentsA4 += 1
		}

		if strings.HasPrefix(validatedLists[k][3], "2024A5") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			A5total += float64(floatConvertScore)
			numberOfStudentsA5 += 1
		}

		if strings.HasPrefix(validatedLists[k][3], "2024A7") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			A7total += float64(floatConvertScore)
			numberOfStudentsA7 += 1
		}
		if strings.HasPrefix(validatedLists[k][3], "2024A8") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			A8total += float64(floatConvertScore)
			numberOfStudentsA8 += 1
		}
		if strings.HasPrefix(validatedLists[k][3], "2024AD") {
			floatConvertScore, err := strconv.ParseFloat(validatedLists[k][10], 64)
			if err != nil {
				log.Fatal(err)
			}
			ADtotal += float64(floatConvertScore)
			numberOfStudentsAD += 1
		}

		quizScore += valueOfQuizMarks
		midsemScore += valueOfmidsemMarks
		labtestScore += valueOflabtestMarks
		weeklyScore += valueOfweeklyMarks
		prectScore += valueOfPreCTMarks
		compreScore += valueOfCompmreMarks
		totalScore += valueOfTotalMarks

	}
	numberOfValidatedLists := float64(len(validatedLists))

	if err != nil {
		log.Fatal(err)
	}

	averageQuizScore := quizScore / numberOfValidatedLists
	averageMidSemScore := midsemScore / numberOfValidatedLists
	averageLabTestScore := labtestScore / numberOfValidatedLists
	averageweeklyScore := weeklyScore / numberOfValidatedLists
	averageprectScore := prectScore / numberOfValidatedLists
	averagecompreScore := compreScore / numberOfValidatedLists
	averagetotalScore := totalScore / numberOfValidatedLists
	averageA3 := A3total / float64(numberOfStudentsA3)
	averageA4 := A4total / float64(numberOfStudentsA4)
	averageA5 := A5total / float64(numberOfStudentsA5)
	averageA7 := A7total / float64(numberOfStudentsA7)
	averageA8 := A8total / float64(numberOfStudentsA8)
	averageAD := ADtotal / float64(numberOfStudentsAD)

	fmt.Println("\n\nAverage Quiz score", averageQuizScore)
	fmt.Println("Average Mid sem score", averageMidSemScore)
	fmt.Println("Average Lab Test score", averageLabTestScore)
	fmt.Println("Average Weekly Lab score", averageweeklyScore)
	fmt.Println("Average PreCT score", averageprectScore)
	fmt.Println("Average Compre score", averagecompreScore)
	fmt.Println("Average Total score", averagetotalScore)

	fmt.Println("\n\nNow printing branch wise averages")

	// Branch-wise average

	fmt.Println("Average score of A3", averageA3)
	fmt.Println("Average score of A4", averageA4)
	fmt.Println("Average score of A5", averageA5)
	fmt.Println("Average score of A7", averageA7)
	fmt.Println("Average score of A8", averageA8)
	fmt.Println("Average score of AD", averageAD)

	// RANK FINDER HAHAHAHA

	// QUIZ RANKS

	sort.Slice(validatedLists, func(i, j int) bool {
		o, err := strconv.ParseFloat(validatedLists[i][4], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][4], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Quiz:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][4]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// MID SEM RANKS

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][5], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][5], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Mid sems:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][5]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// LAB TEST MARKS

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][6], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][6], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Lab tests:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][6]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// Weekly Ranks

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][7], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][7], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Weekly labs:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][7]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// Pre CT Ranks

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][8], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][8], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of PRE CT:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][8]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// Compre Ranks

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][9], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][9], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Compre:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][9]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}

	// Total Ranks

	sort.Slice(validatedLists, func(i, j int) bool {

		o, err := strconv.ParseFloat(validatedLists[i][10], 64)
		if err != nil {
			log.Fatal(err)
		}
		p, err := strconv.ParseFloat(validatedLists[j][10], 64)
		if err != nil {
			log.Fatal(err)
		}
		return o > p
	})
	fmt.Println("\nTop 3 students Of Total marks:")
	for i := 0; i < 3 && i < len(validatedLists); i++ {
		id := validatedLists[i][2]
		rank := i + 1
		marks := validatedLists[i][10]
		fmt.Printf("EMPID: %s, Rank: %d, Marks: %s\n", id, rank, marks)
	}
}
