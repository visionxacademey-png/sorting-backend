const express = require("express");
const cors = require("cors");
const multer = require("multer");
const xlsx = require("xlsx");
const ExcelJS = require("exceljs");

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer({ dest: "uploads/" });

// Time slots
const timeSlots = [
  "9:00-10:30",
  "10:30-12:00",
  "12:00-1:30",
  "1:30-3:00",
  "3:00-4:30",
  "4:30-6:00",
];

// Create class
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
app.post("/upload", upload.single("file"), async (req, res) => {
  try {
    // =====================
    // READ EXCEL
    // =====================
    const inputWorkbook = xlsx.readFile(req.file.path);
    const sheet = inputWorkbook.Sheets[inputWorkbook.SheetNames[0]];

    const students = xlsx.utils.sheet_to_json(sheet).map((s) => ({
      name: s["Student Name"] || s["Name"],
      college: s["College"],
      internship: s["Internship"],
    }));

    // =====================
    // CLASS INPUT FROM FRONTEND
    // =====================
    const classInput = JSON.parse(req.body.classes || "[]");

    let classes = classInput.map((c) =>
      createClass(c.name, Number(c.capacity))
    );

    if (classes.length === 0) {
      return res.status(400).json({ error: "No classes provided" });
    }

    // =====================
    // GROUPING
    // =====================
    const groups = {};

    students.forEach((s) => {
      const key = `${s.college}__${s.internship}`;

      if (!groups[key]) {
        groups[key] = {
          college: s.college,
          internship: s.internship,
          students: [],
        };
      }

      groups[key].students.push(s);
    });

    const groupList = Object.values(groups);

    // =====================
    // INTERNSHIP PRIORITY
    // =====================
    const internshipMap = {};

    groupList.forEach((g) => {
      if (!internshipMap[g.internship]) {
        internshipMap[g.internship] = [];
      }
      internshipMap[g.internship].push(g);
    });

    const internships = Object.keys(internshipMap).sort((a, b) => {
      const totalA = internshipMap[a].reduce((sum, g) => sum + g.students.length, 0);
      const totalB = internshipMap[b].reduce((sum, g) => sum + g.students.length, 0);
      return totalB - totalA;
    });

    let tempClassCount = 1;
    const finalAssignments = [];

    // =====================
    // ALLOCATION (COMBINED LOGIC)
    // =====================
    internships.forEach((internship) => {
      const sortedGroups = internshipMap[internship].sort(
        (a, b) => b.students.length - a.students.length
      );

      sortedGroups.forEach((group) => {
        let bestClass = null;
        let bestSlot = null;
        let minWaste = Infinity;

        for (let cls of classes) {
          for (let slot of cls.slots) {

            // CASE 1: Empty slot
            if (!slot.assigned.internship) {
              if (cls.capacity >= group.students.length) {
                const waste = cls.capacity - group.students.length;

                if (waste < minWaste) {
                  minWaste = waste;
                  bestClass = cls;
                  bestSlot = slot;
                }
              }
            }

            // CASE 2: Same internship → combine
            else if (slot.assigned.internship === internship) {
              const newTotal = slot.assigned.used + group.students.length;

              if (newTotal <= cls.capacity) {
                const waste = cls.capacity - newTotal;

                if (waste < minWaste) {
                  minWaste = waste;
                  bestClass = cls;
                  bestSlot = slot;
                }
              }
            }
          }
        }

        // ASSIGN
        if (bestClass && bestSlot) {

          if (!bestSlot.assigned.internship) {
            bestSlot.assigned.internship = internship;
          }

          bestSlot.assigned.groups.push({
            college: group.college,
            count: group.students.length,
            students: group.students,
          });

          bestSlot.assigned.used += group.students.length;

          group.students.forEach((s) => {
            finalAssignments.push({
              name: s.name,
              college: s.college,
              internship: s.internship,
              class: bestClass.name,
              time: bestSlot.time,
            });
          });
        }

        // TEMP CLASS
        else {
          const tempClassName = `TEMP_${tempClassCount++}`;
          const newClass = createClass(tempClassName, group.students.length);

          newClass.slots[0].assigned.internship = internship;
          newClass.slots[0].assigned.groups.push({
            college: group.college,
            count: group.students.length,
            students: group.students,
          });

          newClass.slots[0].assigned.used = group.students.length;

          group.students.forEach((s) => {
            finalAssignments.push({
              name: s.name,
              college: s.college,
              internship: s.internship,
              class: tempClassName,
              time: newClass.slots[0].time,
            });
          });

          classes.push(newClass);
        }
      });
    });

    // =====================
    // EXCEL EXPORT
    // =====================
    const workbook = new ExcelJS.Workbook();

    // Sheet 1
    const sheet1 = workbook.addWorksheet("Student Allocation");
    sheet1.columns = [
      { header: "Student Name", key: "name", width: 25 },
      { header: "College", key: "college", width: 25 },
      { header: "Internship", key: "internship", width: 20 },
      { header: "Class", key: "class", width: 15 },
      { header: "Time Slot", key: "time", width: 20 },
    ];

    finalAssignments.forEach((s) => {
      sheet1.addRow(s);
    });

    // Sheet 2
    const sheet2 = workbook.addWorksheet("Class Schedule");
    sheet2.columns = [
      { header: "Class", key: "class", width: 15 },
      { header: "Time", key: "time", width: 20 },
      { header: "Internship", key: "internship", width: 20 },
      { header: "Colleges", key: "college", width: 30 },
      { header: "Total Students", key: "count", width: 20 },
    ];

    classes.forEach((cls) => {
      cls.slots.forEach((slot) => {
        if (slot.assigned.internship) {
          sheet2.addRow({
            class: cls.name,
            time: slot.time,
            internship: slot.assigned.internship,
            college: slot.assigned.groups.map(g => g.college).join(", "),
            count: slot.assigned.used,
          });
        }
      });
    });

    // Sheet 3
    const sheet3 = workbook.addWorksheet("Timetable");

    sheet3.addRow([
      "Class",
      "9:00",
      "10:30",
      "12:00",
      "1:30",
      "3:00",
      "4:30",
    ]);

    classes.forEach((cls) => {
      const row = [cls.name];

      cls.slots.forEach((slot) => {
        if (slot.assigned.internship) {
          row.push(
            `${slot.assigned.internship}\n${slot.assigned.groups.map(g => g.college).join(", ")} (${slot.assigned.used})`
          );
        } else {
          row.push("Empty");
        }
      });

      sheet3.addRow(row);
    });

    // SEND FILE
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=allocation.xlsx"
    );

    await workbook.xlsx.write(res);
    res.end();

  } catch (err) {
    console.error(err);
    res.status(500).json({ error: "Error processing file" });
  }
});

// =====================
app.listen(5000, () => {
  console.log("Server running on port 5000 🚀");
});