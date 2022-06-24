const User = [
  {
    fname: "Amir",
    lname: "Mustafa",
    email: "amir@gmail.com",
    gender: "Male",
  },
  {
    fname: "Ashwani",
    lname: "Kumar",
    email: "ashwani@gmail.com",
    gender: "Male",
  },
  {
    fname: "Nupur",
    lname: "Shah",
    email: "nupur@gmail.com",
    gender: "Female",
  },
  {
    fname: "Himanshu",
    lname: "Mewari",
    email: "himanshu@gmail.com",
    gender: "Male",
  },
  {
    fname: "Vankayala",
    lname: "Sirisha",
    email: "sirisha@gmail.com",
    gender: "Female",
  },
];

import excelJS from "exceljs";

const handler = (req, res) => {
  const workbook = new excelJS.Workbook();
  const worksheet = workbook.addWorksheet("My Users");

  worksheet.columns = [
    { header: "S no.", key: "s_no", width: 10 },
    { header: "First Name", key: "fname", width: 10 },
    { header: "Last Name", key: "lname", width: 10 },
    { header: "Email Id", key: "email", width: 10 },
    { header: "Gender", key: "gender", width: 10 },
  ];

  let counter = 1;
  User.forEach((user) => {
    user.s_no = counter;
    worksheet.addRow(user); // Add data in worksheet
    counter++;
  });

  try {
    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=" + "tutorials.xlsx"
    );

    return workbook.xlsx.write(res).then(() => {
      res.status(200).end();
    });
  } catch (err) {
    res.send({
      status: "error",
      message: "Something went wrong",
    });
  }
};

export default handler;
