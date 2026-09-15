const express = require("express");
const cors = require("cors");
const XLSX = require("xlsx");
const path = require("path");
const fs = require("fs");

const app = express();
const PORT = 3000;


// =====================================================
// MIDDLEWARE
// =====================================================

app.use(cors());
app.use(express.json());
app.use(express.static(__dirname));


// =====================================================
// EXCEL FILES
// =====================================================

const STUDENTS_FILE =
    path.join(__dirname, "students.xlsx");

const QUESTIONS_FILE =
    path.join(__dirname, "questions.xlsx");

const RESULTS_FILE =
    path.join(__dirname, "quiz_results.xlsx");

const EVENTS_FILE =
    path.join(__dirname, "events.xlsx");

const EVENT_REGISTRATIONS_FILE =
    path.join(__dirname, "event_registrations.xlsx");


// =====================================================
// CHECK FILE
// =====================================================

function checkFile(file) {

    if (!fs.existsSync(file)) {

        console.log("Missing file:", file);

        return false;
    }

    return true;
}


// =====================================================
// READ EXCEL
// =====================================================

function readExcel(file) {

    if (!fs.existsSync(file)) {
        return [];
    }

    const workbook = XLSX.readFile(file);

    if (!workbook.SheetNames.length) {
        return [];
    }

    const sheetName =
        workbook.SheetNames[0];

    const worksheet =
        workbook.Sheets[sheetName];

    return XLSX.utils.sheet_to_json(
        worksheet,
        {
            defval: ""
        }
    );
}


// =====================================================
// WRITE EXCEL
// =====================================================

function writeExcel(
    file,
    data,
    sheetName
) {

    const workbook =
        XLSX.utils.book_new();

    const worksheet =
        XLSX.utils.json_to_sheet(data);

    XLSX.utils.book_append_sheet(
        workbook,
        worksheet,
        sheetName
    );

    try {

        XLSX.writeFile(
            workbook,
            file
        );

    } catch (error) {

        console.error(
            "Excel write error:",
            error
        );

        if (
            error.code === "EBUSY" ||
            error.code === "EPERM" ||
            error.code === "EACCES"
        ) {

            throw new Error(
                "Cannot update " +
                path.basename(file) +
                ". Please close the Excel file if it is open and try again."
            );
        }

        throw error;
    }
}


// =====================================================
// CREATE QUIZ RESULTS FILE
// =====================================================

function createResultsFile() {

    if (
        !fs.existsSync(
            RESULTS_FILE
        )
    ) {

        const initialData = [
            {
                RegNo: "",
                Date: "",
                QNo: "",
                Answer: "",
                CorrectAnswer: "",
                Score: "",
                Attempted: ""
            }
        ];

        writeExcel(
            RESULTS_FILE,
            initialData,
            "QuizResults"
        );

        console.log(
            "quiz_results.xlsx created."
        );
    }
}

createResultsFile();


// =====================================================
// CREATE EVENT REGISTRATIONS FILE
// =====================================================

function createEventRegistrationsFile() {

    if (
        !fs.existsSync(
            EVENT_REGISTRATIONS_FILE
        )
    ) {

        const initialData = [
            {
                RegNo: "",
                StudentName: "",
                Email: "",
                EventName: "",
                EventDate: "",
                Venue: "",
                RegisteredOn: ""
            }
        ];

        writeExcel(
            EVENT_REGISTRATIONS_FILE,
            initialData,
            "EventRegistrations"
        );

        console.log(
            "event_registrations.xlsx created."
        );
    }
}

createEventRegistrationsFile();


// =====================================================
// TODAY'S DATE
// =====================================================

function getTodayDate() {

    const now = new Date();

    const year =
        now.getFullYear();

    const month =
        String(
            now.getMonth() + 1
        ).padStart(2, "0");

    const day =
        String(
            now.getDate()
        ).padStart(2, "0");

    return (
        `${year}-${month}-${day}`
    );
}


// =====================================================
// CONVERT EXCEL DATE
// =====================================================

function formatExcelDate(value) {

    if (
        value === null ||
        value === undefined ||
        value === ""
    ) {

        return "";
    }


    // Excel serial date
    if (
        typeof value === "number"
    ) {

        const parsed =
            XLSX.SSF.parse_date_code(
                value
            );

        if (!parsed) {
            return "";
        }

        return (
            String(parsed.y) +
            "-" +
            String(parsed.m).padStart(2, "0") +
            "-" +
            String(parsed.d).padStart(2, "0")
        );
    }


    // JavaScript Date
    if (
        value instanceof Date
    ) {

        return (
            value.getFullYear() +
            "-" +
            String(
                value.getMonth() + 1
            ).padStart(2, "0") +
            "-" +
            String(
                value.getDate()
            ).padStart(2, "0")
        );
    }


    const text =
        String(value).trim();


    // DD-MM-YYYY or DD/MM/YYYY
    let match =
        text.match(
            /^(\d{1,2})[-\/](\d{1,2})[-\/](\d{4})$/
        );

    if (match) {

        const day =
            match[1].padStart(2, "0");

        const month =
            match[2].padStart(2, "0");

        const year =
            match[3];

        return (
            `${year}-${month}-${day}`
        );
    }


    // YYYY-MM-DD or YYYY/MM/DD
    match =
        text.match(
            /^(\d{4})[-\/](\d{1,2})[-\/](\d{1,2})$/
        );

    if (match) {

        const year =
            match[1];

        const month =
            match[2].padStart(2, "0");

        const day =
            match[3].padStart(2, "0");

        return (
            `${year}-${month}-${day}`
        );
    }


    return text;
}


// =====================================================
// GET CORRECT ANSWER
// =====================================================

function getCorrectAnswerLetter(
    question
) {

    const answer =
        String(
            question.Answer || ""
        )
        .trim()
        .toLowerCase();


    // Excel contains A/B/C/D
    if (
        answer === "a" ||
        answer === "b" ||
        answer === "c" ||
        answer === "d"
    ) {

        return answer.toUpperCase();
    }


    // Excel contains complete option text
    const options = {

        A:
            String(
                question.OptionA || ""
            )
            .trim()
            .toLowerCase(),

        B:
            String(
                question.OptionB || ""
            )
            .trim()
            .toLowerCase(),

        C:
            String(
                question.OptionC || ""
            )
            .trim()
            .toLowerCase(),

        D:
            String(
                question.OptionD || ""
            )
            .trim()
            .toLowerCase()
    };


    for (
        const letter of
        ["A", "B", "C", "D"]
    ) {

        if (
            options[letter] &&
            options[letter] === answer
        ) {

            return letter;
        }
    }


    return String(
        question.Answer || ""
    ).trim();
}


// =====================================================
// LOGIN
// =====================================================

app.post(
    "/api/login",
    (req, res) => {

        try {

            const {
                regNo,
                password
            } = req.body;


            if (
                !regNo ||
                !password
            ) {

                return res.status(400).json({

                    success: false,

                    message:
                        "Register Number and Password are required."
                });
            }


            if (
                !checkFile(
                    STUDENTS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "students.xlsx was not found."
                });
            }


            const students =
                readExcel(
                    STUDENTS_FILE
                );


            const student =
                students.find(
                    s =>

                        String(
                            s.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()

                        &&

                        String(
                            s.Password
                        ).trim()
                        ===
                        String(
                            password
                        )
                );


            if (!student) {

                return res.status(401).json({

                    success: false,

                    message:
                        "Invalid Register Number or Password."
                });
            }


            // Never send password to browser

            const safeStudent = {
                ...student
            };

            delete safeStudent.Password;


            res.json({

                success: true,

                student:
                    safeStudent
            });


        } catch (error) {

            console.error(
                "Login error:",
                error
            );

            res.status(500).json({

                success: false,

                message:
                    "Unable to login."
            });
        }
    }
);


// =====================================================
// CHANGE PASSWORD
// =====================================================

app.post(
    "/api/change-password",
    (req, res) => {

        try {

            const {
                regNo,
                oldPassword,
                newPassword
            } = req.body;


            if (
                !regNo ||
                !oldPassword ||
                !newPassword
            ) {

                return res.status(400).json({

                    success: false,

                    message:
                        "All fields are required."
                });
            }


            if (
                String(
                    newPassword
                ).length < 6
            ) {

                return res.status(400).json({

                    success: false,

                    message:
                        "New password must be at least 6 characters."
                });
            }


            if (
                !checkFile(
                    STUDENTS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "students.xlsx was not found."
                });
            }


            const workbook =
                XLSX.readFile(
                    STUDENTS_FILE
                );


            const sheetName =
                workbook.SheetNames[0];


            const worksheet =
                workbook.Sheets[
                    sheetName
                ];


            const students =
                XLSX.utils.sheet_to_json(
                    worksheet,
                    {
                        defval: ""
                    }
                );


            const studentIndex =
                students.findIndex(
                    student =>

                        String(
                            student.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()
                );


            if (
                studentIndex === -1
            ) {

                return res.status(404).json({

                    success: false,

                    message:
                        "Student not found."
                });
            }


            const student =
                students[
                    studentIndex
                ];


            if (
                String(
                    student.Password
                ).trim()
                !==
                String(
                    oldPassword
                )
            ) {

                return res.status(401).json({

                    success: false,

                    message:
                        "Old password is incorrect."
                });
            }


            students[
                studentIndex
            ].Password =
                String(newPassword);


            const newWorksheet =
                XLSX.utils.json_to_sheet(
                    students
                );


            workbook.Sheets[
                sheetName
            ] =
                newWorksheet;


            try {

                XLSX.writeFile(
                    workbook,
                    STUDENTS_FILE
                );

            } catch (writeError) {

                if (
                    writeError.code === "EBUSY" ||
                    writeError.code === "EPERM" ||
                    writeError.code === "EACCES"
                ) {

                    throw new Error(
                        "Cannot update students.xlsx. Please close the Excel file and try again."
                    );
                }

                throw writeError;
            }


            console.log(
                `Password changed for ${student.RegNo}`
            );


            res.json({

                success: true,

                message:
                    "Password changed successfully."
            });


        } catch (error) {

            console.error(
                "Password change error:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    error.message ||
                    "Unable to change password."
            });
        }
    }
);


// =====================================================
// GET TODAY'S QUESTION
// =====================================================

app.get(
    "/api/today-question",
    (req, res) => {

        try {

            if (
                !checkFile(
                    QUESTIONS_FILE
                )
            ) {

                return res.status(500).json({

                    available: false,

                    message:
                        "questions.xlsx was not found."
                });
            }


            const today =
                getTodayDate();


            const questions =
                readExcel(
                    QUESTIONS_FILE
                );


            if (
                !questions.length
            ) {

                return res.status(404).json({

                    available: false,

                    message:
                        "questions.xlsx is empty."
                });
            }


            const question =
                questions.find(
                    q =>

                        formatExcelDate(
                            q.Date
                        )
                        ===
                        today
                );


            if (!question) {

                return res.json({

                    available: false,

                    message:
                        "No quiz available today."
                });
            }


            console.log(
                "Today's Question:",
                question.QNo
            );


            // Correct answer is NOT sent
            res.json({

                available: true,

                question: {

                    QNo:
                        question.QNo,

                    Question:
                        question.Question,

                    OptionA:
                        question.OptionA,

                    OptionB:
                        question.OptionB,

                    OptionC:
                        question.OptionC,

                    OptionD:
                        question.OptionD
                }
            });


        } catch (error) {

            console.error(
                "Today's question error:",
                error
            );


            res.status(500).json({

                available: false,

                message:
                    "Unable to load today's question."
            });
        }
    }
);


// =====================================================
// QUIZ STATUS
// =====================================================

app.get(
    "/api/quiz-status/:regNo",
    (req, res) => {

        try {

            const regNo =
                req.params.regNo;

            const today =
                getTodayDate();


            if (
                !checkFile(
                    RESULTS_FILE
                )
            ) {

                return res.json({

                    attempted: false
                });
            }


            const results =
                readExcel(
                    RESULTS_FILE
                );


            const attempt =
                results.find(
                    result =>

                        String(
                            result.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()

                        &&

                        formatExcelDate(
                            result.Date
                        )
                        ===
                        today

                        &&

                        String(
                            result.Attempted
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        "yes"
                );


            if (attempt) {

                return res.json({

                    attempted: true,

                    QNo:
                        attempt.QNo
                });
            }


            res.json({

                attempted: false
            });


        } catch (error) {

            console.error(
                "Quiz status error:",
                error
            );


            res.status(500).json({

                attempted: false,

                message:
                    "Unable to check quiz status."
            });
        }
    }
);


// =====================================================
// SUBMIT QUIZ
// =====================================================

app.post(
    "/api/submit-quiz",
    (req, res) => {

        try {

            const {
                regNo,
                QNo,
                answer
            } = req.body;


            if (
                !regNo ||
                QNo === undefined ||
                !answer
            ) {

                return res.status(400).json({

                    success: false,

                    message:
                        "Quiz information is missing."
                });
            }


            const today =
                getTodayDate();


            // ---------------------------------------------
            // READ QUESTIONS
            // ---------------------------------------------

            if (
                !checkFile(
                    QUESTIONS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "questions.xlsx was not found."
                });
            }


            const questions =
                readExcel(
                    QUESTIONS_FILE
                );


            const question =
                questions.find(
                    q =>

                        formatExcelDate(
                            q.Date
                        )
                        ===
                        today

                        &&

                        String(
                            q.QNo
                        ).trim()
                        ===
                        String(
                            QNo
                        ).trim()
                );


            if (!question) {

                return res.status(404).json({

                    success: false,

                    message:
                        "Today's question was not found."
                });
            }


            // ---------------------------------------------
            // READ RESULTS
            // ---------------------------------------------

            let results =
                readExcel(
                    RESULTS_FILE
                );


            // Remove blank starter row

            results =
                results.filter(
                    r =>

                        String(
                            r.RegNo || ""
                        ).trim()
                        !==
                        ""
                );


            // ---------------------------------------------
            // PREVENT SECOND ATTEMPT
            // ---------------------------------------------

            const existingAttempt =
                results.find(
                    result =>

                        String(
                            result.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()

                        &&

                        formatExcelDate(
                            result.Date
                        )
                        ===
                        today

                        &&

                        String(
                            result.Attempted
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        "yes"
                );


            if (existingAttempt) {

                return res.status(409).json({

                    success: false,

                    message:
                        "You have already attempted today's quiz."
                });
            }


            // ---------------------------------------------
            // CHECK ANSWER
            // ---------------------------------------------

            const correctAnswer =
                getCorrectAnswerLetter(
                    question
                );


            const studentAnswer =
                String(
                    answer
                )
                .trim()
                .toUpperCase();


            const isCorrect =
                studentAnswer ===
                String(
                    correctAnswer
                )
                .trim()
                .toUpperCase();


            const score =
                isCorrect
                    ? 1
                    : 0;


            const reason =
                String(
                    question.Reason || ""
                ).trim();


            // ---------------------------------------------
            // SAVE QUIZ RESULT
            // ---------------------------------------------

            results.push({

                RegNo:
                    String(regNo).trim(),

                Date:
                    today,

                QNo:
                    QNo,

                Answer:
                    studentAnswer,

                CorrectAnswer:
                    correctAnswer,

                Score:
                    score,

                Attempted:
                    "Yes"
            });


            writeExcel(
                RESULTS_FILE,
                results,
                "QuizResults"
            );


            // ---------------------------------------------
            // UPDATE STUDENT SCORE
            // ---------------------------------------------

            const students =
                readExcel(
                    STUDENTS_FILE
                );


            const studentIndex =
                students.findIndex(
                    student =>

                        String(
                            student.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()
                );


            let totalScore =
                score;


            if (
                studentIndex !== -1
            ) {

                let oldScore =
                    Number(
                        students[
                            studentIndex
                        ].QuizScore
                    );


                if (
                    isNaN(oldScore)
                ) {

                    oldScore = 0;
                }


                totalScore =
                    oldScore +
                    score;


                students[
                    studentIndex
                ].QuizScore =
                    totalScore;


                students[
                    studentIndex
                ].QuizAttempted =
                    "Yes";


                writeExcel(
                    STUDENTS_FILE,
                    students,
                    "Students"
                );
            }


            res.json({

                success: true,

                correct:
                    isCorrect,

                totalScore:
                    totalScore,

                reason:
                    reason,

                message:
                    isCorrect
                        ? "Correct answer!"
                        : "Incorrect answer."
            });


        } catch (error) {

            console.error(
                "Quiz submission error:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    error.message ||
                    "Unable to submit quiz."
            });
        }
    }
);


// =====================================================
// GET QUIZ RESULTS FOR ONE STUDENT
// =====================================================

app.get(
    "/api/quiz-results/:regNo",
    (req, res) => {

        try {

            const regNo =
                req.params.regNo;


            const results =
                readExcel(
                    RESULTS_FILE
                );


            const studentResults =
                results.filter(
                    result =>

                        String(
                            result.RegNo
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        String(
                            regNo
                        )
                        .trim()
                        .toLowerCase()
                );


            res.json({

                success: true,

                results:
                    studentResults
            });


        } catch (error) {

            console.error(
                "Quiz results error:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    "Unable to load quiz results."
            });
        }
    }
);


// =====================================================
// GET EVENTS
// =====================================================
//
// Excel columns:
//
// Name of the event
// Date
// Venue
// Description
// Poster
// Display
//
// Only Display = Yes events are shown.
// =====================================================

app.get(
    "/api/events",
    (req, res) => {

        try {

            if (
                !checkFile(
                    EVENTS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "events.xlsx was not found."
                });
            }


            const events =
                readExcel(
                    EVENTS_FILE
                );


            const visibleEvents =
                events.filter(
                    event =>

                        String(
                            event.Display || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        "yes"
                );


            const formattedEvents =
                visibleEvents.map(
                    event => ({

                        eventName:
                            String(
                                event["Name of the event"] || ""
                            ).trim(),

                        date:
                            formatExcelDate(
                                event.Date
                            ),

                        venue:
                            String(
                                event.Venue || ""
                            ).trim(),

                        description:
                            String(
                                event.Description || ""
                            ).trim(),

                        poster:
                            String(
                                event.Poster || ""
                            ).trim(),

                        display:
                            true
                    })
                );


            console.log(
                `Visible events: ${formattedEvents.length}`
            );


            res.json({

                success: true,

                events:
                    formattedEvents
            });


        } catch (error) {

            console.error(
                "Events error:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    "Unable to load events."
            });
        }
    }
);


// =====================================================
// REGISTER FOR EVENT
// =====================================================

app.post(
    "/api/event-register",
    (req, res) => {

        try {

            const {
                regNo,
                eventName
            } = req.body;


            // ---------------------------------------------
            // VALIDATE INPUT
            // ---------------------------------------------

            if (
                !regNo ||
                !eventName
            ) {

                return res.status(400).json({

                    success: false,

                    message:
                        "Register Number and Event Name are required."
                });
            }


            const cleanRegNo =
                String(
                    regNo
                ).trim();


            const cleanEventName =
                String(
                    eventName
                ).trim();


            // ---------------------------------------------
            // CHECK STUDENTS FILE
            // ---------------------------------------------

            if (
                !checkFile(
                    STUDENTS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "students.xlsx was not found."
                });
            }


            // ---------------------------------------------
            // FIND STUDENT
            // ---------------------------------------------

            const students =
                readExcel(
                    STUDENTS_FILE
                );


            const student =
                students.find(
                    s =>

                        String(
                            s.RegNo || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        cleanRegNo
                            .toLowerCase()
                );


            if (!student) {

                return res.status(404).json({

                    success: false,

                    message:
                        "Student not found."
                });
            }


            // ---------------------------------------------
            // CHECK EVENTS FILE
            // ---------------------------------------------

            if (
                !checkFile(
                    EVENTS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "events.xlsx was not found."
                });
            }


            // ---------------------------------------------
            // FIND EVENT
            // ---------------------------------------------

            const events =
                readExcel(
                    EVENTS_FILE
                );


            const event =
                events.find(
                    e =>

                        String(
                            e["Name of the event"] || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        cleanEventName.toLowerCase()

                        &&

                        String(
                            e.Display || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        "yes"
                );


            if (!event) {

                return res.status(404).json({

                    success: false,

                    message:
                        "This event is not available for registration."
                });
            }


            // ---------------------------------------------
            // CHECK REGISTRATION FILE
            // ---------------------------------------------

            if (
                !checkFile(
                    EVENT_REGISTRATIONS_FILE
                )
            ) {

                return res.status(500).json({

                    success: false,

                    message:
                        "event_registrations.xlsx was not found."
                });
            }


            // ---------------------------------------------
            // READ REGISTRATIONS
            // ---------------------------------------------

            let registrations =
                readExcel(
                    EVENT_REGISTRATIONS_FILE
                );


            // Remove blank starter rows

            registrations =
                registrations.filter(
                    registration =>

                        String(
                            registration.RegNo || ""
                        ).trim()
                        !==
                        ""
                );


            // ---------------------------------------------
            // CHECK DUPLICATE
            // ---------------------------------------------

            const alreadyRegistered =
                registrations.find(
                    registration =>

                        String(
                            registration.RegNo || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        cleanRegNo.toLowerCase()

                        &&

                        String(
                            registration.EventName || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        cleanEventName.toLowerCase()
                );


            if (alreadyRegistered) {

                return res.status(409).json({

                    success: false,

                    alreadyRegistered: true,

                    message:
                        "You have already registered for this event."
                });
            }


            // ---------------------------------------------
            // CREATE REGISTRATION
            // ---------------------------------------------

            const registration = {

                RegNo:
                    String(
                        student.RegNo
                    ).trim(),

                StudentName:
                    String(
                        student.Name || ""
                    ).trim(),

                Email:
                    String(
                        student.Email || ""
                    ).trim(),

                EventName:
                    String(
                        event["Name of the event"] || ""
                    ).trim(),

                EventDate:
                    formatExcelDate(
                        event.Date
                    ),

                Venue:
                    String(
                        event.Venue || ""
                    ).trim(),

                RegisteredOn:
                    getTodayDate()
            };


            // ---------------------------------------------
            // ADD TO ARRAY
            // ---------------------------------------------

            registrations.push(
                registration
            );


            // ---------------------------------------------
            // WRITE TO EXCEL
            // ---------------------------------------------

            writeExcel(
                EVENT_REGISTRATIONS_FILE,
                registrations,
                "EventRegistrations"
            );


            console.log(
                "======================================"
            );

            console.log(
                "EVENT REGISTRATION SUCCESS"
            );

            console.log(
                "Student:",
                registration.RegNo
            );

            console.log(
                "Event:",
                registration.EventName
            );

            console.log(
                "======================================"
            );


            // ---------------------------------------------
            // SEND SUCCESS
            // ---------------------------------------------

            return res.json({

                success: true,

                message:
                    "Successfully registered for the event.",

                registration:
                    registration
            });


        } catch (error) {

            console.error(
                "======================================"
            );

            console.error(
                "EVENT REGISTRATION ERROR"
            );

            console.error(
                error
            );

            console.error(
                "======================================"
            );


            return res.status(500).json({

                success: false,

                message:
                    error.message ||
                    "Unable to register for the event."
            });
        }
    }
);


// =====================================================
// GET REGISTERED EVENTS FOR ONE STUDENT
// =====================================================

app.get(
    "/api/event-registrations/:regNo",
    (req, res) => {

        try {

            const regNo =
                String(
                    req.params.regNo
                ).trim();


            if (
                !checkFile(
                    EVENT_REGISTRATIONS_FILE
                )
            ) {

                return res.json({

                    success: true,

                    registrations: []
                });
            }


            const registrations =
                readExcel(
                    EVENT_REGISTRATIONS_FILE
                );


            const studentRegistrations =
                registrations.filter(
                    registration =>

                        String(
                            registration.RegNo || ""
                        )
                        .trim()
                        .toLowerCase()
                        ===
                        regNo.toLowerCase()
                );


            res.json({

                success: true,

                registrations:
                    studentRegistrations
            });


        } catch (error) {

            console.error(
                "Event registrations error:",
                error
            );


            res.status(500).json({

                success: false,

                message:
                    "Unable to load event registrations."
            });
        }
    }
);


// =====================================================
// START SERVER
// =====================================================

app.listen(
    PORT,
    () => {

        console.log("");

        console.log(
            "======================================"
        );

        console.log(
            "        EHC PORTAL IS RUNNING"
        );

        console.log(
            "======================================"
        );

        console.log("");

        console.log(
            `Open: http://localhost:${PORT}`
        );

        console.log("");

        console.log(
            "Students Excel:",
            STUDENTS_FILE
        );

        console.log(
            "Questions Excel:",
            QUESTIONS_FILE
        );

        console.log(
            "Quiz Results Excel:",
            RESULTS_FILE
        );

        console.log(
            "Events Excel:",
            EVENTS_FILE
        );

        console.log(
            "Event Registrations Excel:",
            EVENT_REGISTRATIONS_FILE
        );

        console.log("");
    }
);