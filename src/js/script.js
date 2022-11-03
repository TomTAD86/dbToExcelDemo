// AJAX Function

const getEmployees = async () => {
  const employees = await new Promise((resolve, reject) => {
    $.ajax({
      url: "src/php/getEmployeeData.php",
      type: "GET",
      dataType: "JSON",
      success: function (result) {
        resolve(result);
      },
      error: function (jqXHR, textStatus, errorThrown) {
        console.log(jqXHR);
        console.log(textStatus);
        console.log(JSON.stringify(errorThrown));
        reject(JSON.stringify(errorThrown));
      },
    });
  });
  return await employees.data;
};

$(async function () {
  const employees = await getEmployees();
  employees.forEach((employee) => {
    const newDiv = `
    <tr>
      <td>${employee.firstName}</td>
      <td>${employee.lastName}</td>
      <td>${employee.position}</td>
      <td>${employee.department}</td>
      <td>${employee.salary}</td>
      <td>${employee.salesThisYear}</td>
    </tr>`;
    $("#employee-table").append(newDiv);
  });
  console.log(employees);
});

$("#generate-excel").click(async function() {
  $("#generate-excel").text("Generating...");
  console.log("Generating...");
  await $.ajax({
    url: "src/php/createExcelFile.php",
    type: "GET",
  });
  $("#generate-excel").text("Excel File Generated");
});