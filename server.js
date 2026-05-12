const express = require("express");
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");

const app = express();

app.use(cors({
  origin: "*",
}));

app.use(express.json());

const upload = multer({
  dest: "uploads/",
});

// =====================
// TIME SLOTS
// =====================
const timeSlots = [
  "9:00-10:30",
  "10:30-12:00",
  "12:00-1:30",
  "1:30-3:00",
  "3:00-4:30",
  "4:30-6:00",
];

// =====================
// CREATE CLASS
// =====================
const createClass = (name, capacity) => ({
  name,
  capacity,
  slots: timeSlots.map((time) => ({
    time,
    assigned: {
      internship: null,
      groups: [],
      used: 0,
    },
  })),
});

// =====================
// MAIN API
// =====================
app.post(
  "/upload",
  upload.single("file"),
  async (req, res) => {

    try {

      // =====================
      // READ EXCEL
      // =====================
      const workbook =
        xlsx.readFile(req.file.path);

      const sheet =
        workbook.Sheets[
          workbook.SheetNames[0]
        ];

      // =====================
      // PARSE STUDENTS
      // =====================
      const students =
        xlsx.utils
          .sheet_to_json(sheet)
          .map((row) => ({

            // NAME
            name:
              row["Name"] ||
              row["Student Name"] ||
              "",

            // PHONE
            phone: String(
              row["Phone"] ??
              row["Phone Number"] ??
              ""
            ).trim(),

            // COLLEGE
            college:
              row["College"] ||
              row["College Name"] ||
              "",

            // INTERNSHIP
            internship:
              row["Internship"] ||
              row["Internship "] ||
              row["Internship Name"] ||
              "",

          }));

      console.log(students[0]);

      // =====================
      // CLASS INPUT
      // =====================
      const classInput =
        JSON.parse(
          req.body.classes || "[]"
        );

      let classes =
        classInput.map((c) =>
          createClass(
            c.name,
            Number(c.capacity)
          )
        );

      // =====================
      // GROUPING
      // =====================
      const groups = {};

      students.forEach((s) => {

        const key =
          `${s.college}__${s.internship}`;

        if (!groups[key]) {

          groups[key] = {

            college:
              s.college,

            internship:
              s.internship,

            students: [],

          };

        }

        groups[key]
          .students
          .push(s);

      });

      const groupList =
        Object.values(groups);

      // =====================
      // INTERNSHIP PRIORITY
      // =====================
      const internshipMap = {};

      groupList.forEach((g) => {

        if (
          !internshipMap[
            g.internship
          ]
        ) {

          internshipMap[
            g.internship
          ] = [];

        }

        internshipMap[
          g.internship
        ].push(g);

      });

      const internships =
        Object.keys(
          internshipMap
        ).sort((a, b) => {

          const totalA =
            internshipMap[a]
              .reduce(
                (sum, g) =>
                  sum +
                  g.students.length,
                0
              );

          const totalB =
            internshipMap[b]
              .reduce(
                (sum, g) =>
                  sum +
                  g.students.length,
                0
              );

          return totalB - totalA;

        });

      // =====================
      // ALLOCATION
      // =====================
      let tempClassCount = 1;

      const finalAssignments = [];

      internships.forEach(
        (internship) => {

          const sortedGroups =
            internshipMap[
              internship
            ].sort(
              (a, b) =>
                b.students.length -
                a.students.length
            );

          sortedGroups.forEach(
            (group) => {

              let bestClass =
                null;

              let bestSlot =
                null;

              let minWaste =
                Infinity;

              // =====================
              // FIND BEST SLOT
              // =====================
              for (
                let cls of classes
              ) {

                for (
                  let slot of cls.slots
                ) {

                  // EMPTY SLOT
                  if (
                    !slot.assigned
                      .internship
                  ) {

                    if (
                      cls.capacity >=
                      group.students
                        .length
                    ) {

                      const waste =
                        cls.capacity -
                        group.students
                          .length;

                      if (
                        waste <
                        minWaste
                      ) {

                        minWaste =
                          waste;

                        bestClass =
                          cls;

                        bestSlot =
                          slot;

                      }

                    }

                  }

                  // SAME INTERNSHIP
                  else if (
                    slot.assigned
                      .internship ===
                    internship
                  ) {

                    const newTotal =
                      slot.assigned
                        .used +
                      group.students
                        .length;

                    if (
                      newTotal <=
                      cls.capacity
                    ) {

                      const waste =
                        cls.capacity -
                        newTotal;

                      if (
                        waste <
                        minWaste
                      ) {

                        minWaste =
                          waste;

                        bestClass =
                          cls;

                        bestSlot =
                          slot;

                      }

                    }

                  }

                }

              }

              // =====================
              // ASSIGN
              // =====================
              if (
                bestClass &&
                bestSlot
              ) {

                if (
                  !bestSlot.assigned
                    .internship
                ) {

                  bestSlot.assigned
                    .internship =
                    internship;

                }

                bestSlot.assigned
                  .groups.push({

                    college:
                      group.college,

                    count:
                      group.students
                        .length,

                    students:
                      group.students,

                  });

                bestSlot.assigned
                  .used +=
                  group.students
                    .length;

                group.students
                  .forEach((s) => {

                    finalAssignments
                      .push({

                        name:
                          s.name,

                        phone:
                          s.phone,

                        college:
                          s.college,

                        internship:
                          s.internship,

                        class:
                          bestClass.name,

                        time:
                          bestSlot.time,

                      });

                  });

              }

              // =====================
              // TEMP CLASS
              // =====================
              else {

                const tempClassName =
                  `TEMP_${tempClassCount++}`;

                const newClass =
                  createClass(
                    tempClassName,
                    group.students
                      .length
                  );

                newClass
                  .slots[0]
                  .assigned
                  .internship =
                  internship;

                newClass
                  .slots[0]
                  .assigned
                  .groups
                  .push({

                    college:
                      group.college,

                    count:
                      group.students
                        .length,

                    students:
                      group.students,

                  });

                newClass
                  .slots[0]
                  .assigned
                  .used =
                  group.students
                    .length;

                group.students
                  .forEach((s) => {

                    finalAssignments
                      .push({

                        name:
                          s.name,

                        phone:
                          s.phone,

                        college:
                          s.college,

                        internship:
                          s.internship,

                        class:
                          tempClassName,

                        time:
                          newClass
                            .slots[0]
                            .time,

                      });

                  });

                classes.push(
                  newClass
                );

              }

            }
          );

        }
      );

      // =====================
      // EXPORT EXCEL
      // =====================
      const outputWorkbook =
        new ExcelJS.Workbook();

      // =====================
      // STUDENT ALLOCATION SHEET
      // =====================
      const sheet1 =
        outputWorkbook.addWorksheet(
          "Student Allocation"
        );

      sheet1.columns = [

        {
          header:
            "Student Name",
          key: "name",
          width: 25,
        },

        {
          header:
            "Phone Number",
          key: "phone",
          width: 20,
        },

        {
          header:
            "College",
          key: "college",
          width: 30,
        },

        {
          header:
            "Internship",
          key: "internship",
          width: 25,
        },

        {
          header:
            "Class",
          key: "class",
          width: 15,
        },

        {
          header:
            "Time Slot",
          key: "time",
          width: 20,
        },

      ];

      // =====================
      // ADD STUDENT ROWS
      // =====================
      finalAssignments.forEach(
        (s) => {

          sheet1.addRow({

            name:
              s.name,

            phone:
              s.phone,

            college:
              s.college,

            internship:
              s.internship,

            class:
              s.class,

            time:
              s.time,

          });

        }
      );

      // =====================
      // CLASS SCHEDULE SHEET
      // =====================
      const sheet2 =
        outputWorkbook.addWorksheet(
          "Class Schedule"
        );

      sheet2.columns = [

        {
          header:
            "Class",
          key: "class",
          width: 15,
        },

        {
          header:
            "Time Slot",
          key: "time",
          width: 20,
        },

        {
          header:
            "Internship",
          key: "internship",
          width: 25,
        },

        {
          header:
            "Colleges",
          key: "college",
          width: 40,
        },

        {
          header:
            "Student Count",
          key: "count",
          width: 20,
        },

      ];

      classes.forEach((cls) => {

        cls.slots.forEach(
          (slot) => {

            if (
              slot.assigned
                .internship
            ) {

              sheet2.addRow({

                class:
                  cls.name,

                time:
                  slot.time,

                internship:
                  slot.assigned
                    .internship,

                college:
                  slot.assigned
                    .groups
                    .map(
                      (g) =>
                        g.college
                    )
                    .join(", "),

                count:
                  slot.assigned
                    .used,

              });

            }

          }
        );

      });

      // =====================
      // SEND FILE
      // =====================
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );

      res.setHeader(
        "Content-Disposition",
        "attachment; filename=allocation.xlsx"
      );

      await outputWorkbook
        .xlsx
        .write(res);

      res.end();

    } catch (err) {

      console.error(err);

      res.status(500).json({
        error:
          "Error processing file",
      });

    }

  }
);

// =====================
// SERVER
// =====================
const PORT =
  process.env.PORT || 5000;

app.listen(PORT, () => {

  console.log(
    `Server running on port ${PORT} 🚀`
  );

});