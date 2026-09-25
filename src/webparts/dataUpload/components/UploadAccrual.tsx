import * as React from "react";
import "./UploadAccrual.scss";
import { SPComponentLoader } from "@microsoft/sp-loader";
import { sp } from "../DataUploadWebPart";
import "@pnp/sp/webs";
import "@pnp/sp/site-users/web";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/files";
import "@pnp/sp/folders";
import * as XLSX from "xlsx";
import { SPHttpClient, ISPHttpClientOptions } from "@microsoft/sp-http";

import edit from "../../dataUpload/assets/Pencil.png";
import del from "../../dataUpload/assets/delete.png";

import { IDataUploadProps } from "./IDataUploadProps";

SPComponentLoader.loadCss(
  "https://cdn.jsdelivr.net/npm/bootstrap@5.3.3/dist/css/bootstrap.min.css",
);

export default function UploadAccrual(props: IDataUploadProps) {
  const [file, setFile] = React.useState<File | null>(null);
  const [selectedUser, setSelectedUser] = React.useState<any>(null);
  const [selectedRows, setSelectedRows] = React.useState<number[]>([]);
  const [editingRow, setEditingRow] = React.useState<number | null>(null);
  const [excelData, setExcelData] = React.useState<any[]>([]);
  const [submittedData, setSubmittedData] = React.useState<any[]>([]);
  const [duplicateData, setDuplicateData] = React.useState<any[]>([]);
  const [data, setData] = React.useState<any[]>([]);
  const [filteredData, setFilteredData] = React.useState<any[]>([]);
  const [isSearched, setIsSearched] = React.useState(false);
  const [validationErrors, setValidationErrors] = React.useState<any[]>([]);
  const [errors, setErrors] = React.useState<any[]>([]);
  const [isSubmitting, setIsSubmitting] = React.useState(false);

  const [employee, setEmployee] = React.useState<any>({});
  const normalize = (str: string) => str.trim().toLowerCase();
  let skippedRows: any[] = [];
  const handleRowSelect = (index: number) => {
    setSelectedRows((prev) =>
      prev.includes(index) ? prev.filter((i) => i !== index) : [...prev, index],
    );
  };
  const handleSelectAll = (checked: boolean) => {
    if (checked) {
      setSelectedRows(excelData.map((_, index) => index));
    } else {
      setSelectedRows([]);
    }
  };

  const handleCellChange = (rowIndex: number, field: string, value: string) => {
    const updatedData = [...excelData];

    updatedData[rowIndex] = {
      ...updatedData[rowIndex],
      [field]: value,
    };

    setExcelData(updatedData);
  };

  const loadDuplicateData = async () => {
    const existingItems = await sp.web.lists
      .getByTitle("AccrualSheetList")
      .items();

    console.log(existingItems);
  };
  const saveRow = () => {
    setEditingRow(null);
    alert("Row updated successfully");
  };

  const deleteRow = (index: number) => {
    if (!window.confirm("Delete this row?")) return;

    const updatedData = excelData.filter((_, i) => i !== index);

    setExcelData(updatedData);
  };

  const deleteSelectedRows = () => {
    if (selectedRows.length === 0) {
      alert("Please select rows");
      return;
    }

    if (!window.confirm(`Delete ${selectedRows.length} selected rows?`)) return;

    const updatedData = excelData.filter(
      (_, index) => !selectedRows.includes(index),
    );

    setExcelData(updatedData);
    setSelectedRows([]);

    alert("Rows deleted successfully");
  };

  const requiredColumns = [
    "UserName ",
    "Cost Center ",
    "Cost Center Name ",
    "Vendor Name",
    "Vendor Code ",
    "PO Number",
    "GL Code ",
    "GL Description ",
    "Amount ",
    "Expense Month ",
    "Remarks (if any)",
  ];

  const optionalColumns = ["Remarks (if any)"];

  const validateTemplate = (data: any[]) => {
    if (data.length === 0) {
      alert("Uploaded file is empty");
      return false;
    }

    const fileHeaders = Object.keys(data[0]).map(normalize);
    const allowed = requiredColumns.map(normalize);
    const mandatory = requiredColumns
      .filter((col) => !optionalColumns.includes(col))
      .map(normalize);

    const missing = mandatory.filter((col) => !fileHeaders.includes(col));

    const extra = fileHeaders.filter((col) => !allowed.includes(col));

    if (missing.length > 0) {
      alert("Missing columns: " + missing.join(", "));
      return false;
    }

    if (extra.length > 0) {
      alert("Invalid extra columns: " + extra.join(", "));
      return false;
    }

    return true;
  };
  const handleFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    if (e.target.files && e.target.files.length > 0) {
      const selectedFile = e.target.files[0];
      setFile(selectedFile);

      console.log("Selected file:", selectedFile.name);
    }

    e.target.value = "";
  };

  const ensureUser = async (email: string): Promise<number> => {
    if (!email) return 0;

    try {
      const webUrl = props.context.pageContext.web.absoluteUrl;

      const response = await props.context.spHttpClient.post(
        `${webUrl}/_api/web/ensureuser`,
        SPHttpClient.configurations.v1,
        {
          headers: {
            Accept: "application/json;odata=nometadata",
            "Content-Type": "application/json",
          },
          body: JSON.stringify({
            logonName: email,
          }),
        },
      );

      if (!response.ok) {
        console.log("ensureUser failed for:", email);

        return 0;
      }

      const data = await response.json();

      return data.Id || 0;
    } catch (error) {
      console.log("ensureUser error:", email, error);

      return 0;
    }
  };
  const getuserData = async () => {
    debugger;
    try {
      const toTitleCase = (str: string): string => {
        if (!str) return "";

        return str
          .toLowerCase()
          .split(" ")
          .filter(Boolean)
          .map((w: string) => w.charAt(0).toUpperCase() + w.slice(1))
          .join(" ");
      };

      const cleanLocationForDisplay = (location: string): string => {
        if (!location) return "N/A";

        return location.replace(/^re\s+/i, "").trim();
      };

      const fetchPage = async (pageNumber: number) => {
        debugger;
        const username = "0le867nyvalvfo249e6sj4ri";
        const password =
          "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i";

        const auth = btoa(`${username}:${password}`);

        const responsesevices = await fetch(
          "https://mservices.zinghr.com/etl/api/v2/Auth/GenerateJWTToken?apiPermission=GEMD",
          {
            method: "GET",
            headers: {
              Authorization: `Basic ${auth}`,
              Accept: "application/json",
              "Content-Type": "application/json",
            },
          },
        );

        const ServicedataToken = await responsesevices.json();
        console.log(ServicedataToken.data);

        const response = await fetch(
          "https://mservices.zinghr.com/etl/api/v2/Employee/GetEmployeeDetails",
          {
            method: "POST",
            headers: {
              "Content-Type": "application/json",
              Authorization: `Bearer ${ServicedataToken.data}`,
              ClientSecret:
                "2mpvr7r19amf7o01hr0qncr861hmtsb7o9ap51hwar72405atj3y73mndkmokg5i",
            },
            body: JSON.stringify({
              PageSize: 500,
              PageNumber: pageNumber,
            }),
          },
        );
        if (!response.ok) {
          throw new Error("Failed to fetch employee data");
        }

        return response.json();
      };

      const userEmail = props.context.pageContext.user.email.toLowerCase();

      let item = null;
      let page = 1;

      while (true) {
        const res = await fetchPage(page);

        const employees = res?.data?.employees || [];

        item = employees.find((x: any) => x.email?.toLowerCase() === userEmail);

        if (item) {
          break;
        }

        if (employees.length < 500) {
          break;
        }

        page++;
      }

      if (!item) {
        console.log("Employee not found");
        return;
      }

      const locationAttr = (item.attributes || []).find(
        (a: any) => a.attributeTypeDescription === "Location",
      );

      const departmentAttr = (item.attributes || []).find(
        (a: any) => a.attributeTypeDescription?.toLowerCase() === "department",
      );

      const hodEmailAttr = (item.attributes || []).find(
        (a: any) => a.attributeTypeDescription?.toLowerCase() === "hod_email",
      );

      const hodNameAttr = (item.attributes || []).find(
        (a: any) => a.attributeTypeDescription?.toLowerCase() === "hod name",
      );

      const hodCodeAttr = (item.attributes || []).find(
        (a: any) => a.attributeTypeDescription?.toLowerCase() === "hod_code",
      );

      let employeeUserId = 0;
      let rmUserId = 0;
      let hodUserId = 0;

      try {
        if (item.email) {
          employeeUserId = await ensureUser(item.email);
        }

        if (item.reportingManagerEmail) {
          rmUserId = await ensureUser(item.reportingManagerEmail);
        }

        if (hodEmailAttr?.attributeTypeUnitDescription) {
          hodUserId = await ensureUser(
            hodEmailAttr.attributeTypeUnitDescription,
          );
        }
      } catch (e) {
        console.log("ensureUser error", e);
      }
      debugger;
      console.log(item);
      setEmployee({
        EmployeeCode: item.employeeCode || "",

        EmployeeName: toTitleCase(item.employeeName || ""),

        userEmail: toTitleCase(item.email || ""),

        Division: departmentAttr?.attributeTypeUnitDescription || "",

        Location: cleanLocationForDisplay(
          locationAttr?.attributeTypeUnitDescription || "",
        ),

        RM: item.reportingManagerName || "",

        HOD: hodNameAttr?.attributeTypeUnitDescription || "",

        ContactNo: item.mobileNo || "",

        EmployeeStatus: item.employeeStatus || "",

        Email: item.email || "",

        RMId: rmUserId || 0,

        HODId: hodUserId || 0,
      });

      console.log(hodNameAttr?.attributeTypeUnitDescription);

      const employeeData = {
        EmployeeCode: item.employeeCode || "",
        EmployeeName: toTitleCase(item.employeeName || ""),
        userEmail: toTitleCase(item.email || ""),
        Division: departmentAttr?.attributeTypeUnitDescription || "",
        Location: cleanLocationForDisplay(
          locationAttr?.attributeTypeUnitDescription || "",
        ),
        RM: item.reportingManagerName || "",
        HOD: hodNameAttr?.attributeTypeUnitDescription || "",
        ContactNo: item.mobileNo || "",
        EmployeeStatus: item.employeeStatus || "",
        Email: item.email || "",
        RMId: rmUserId || 0,
        HODId: hodUserId || 0,
      };

      console.log("Before setEmployee:", employeeData);

      setEmployee(employeeData);

      const userApprovers = [rmUserId, hodUserId].filter(
        (id): id is number => !!id,
      );

      const uniqueApprovers = userApprovers.filter(
        (value, index, self) => self.indexOf(value) === index,
      );
    } catch (error) {
      console.error("Error fetching user data:", error);
    }
  };

  React.useEffect(() => {
    void getuserData();
  }, []);

  const downloadTemplate = () => {
    const worksheet = XLSX.utils.json_to_sheet([]);

    XLSX.utils.sheet_add_aoa(worksheet, [requiredColumns]);

    const workbook = {
      Sheets: { Template: worksheet },
      SheetNames: ["Template"],
    };

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    const fileName = "Accrual_Template.xlsx";

    const link = document.createElement("a");
    link.href = URL.createObjectURL(blob);
    link.download = fileName;
    link.click();
  };

  const submitData = async () => {
    if (isSubmitting) {
      return;
    }

    if (excelData.length === 0) {
      alert("No data to submit");
      return;
    }

    setIsSubmitting(true);

    const getMonthNumber = (monthName: string) => {
      const months: any = {
        january: 0,
        february: 1,
        march: 2,
        april: 3,
        may: 4,
        june: 5,
        july: 6,
        august: 7,
        september: 8,
        october: 9,
        november: 10,
        december: 11,
      };

      return months[monthName.trim().toLowerCase()];
    };

    const capitalizeMonth = (month: string) => {
      const m = month.trim().toLowerCase();
      return m.charAt(0).toUpperCase() + m.slice(1);
    };

    const today = new Date();
    const currentMonth = today.getMonth();
    const currentDate = today.getDate();

    let errorList: any[] = [];
    let validRows: any[] = [];

    for (let i = 0; i < excelData.length; i++) {
      const row = excelData[i];
      const rowNumber = i + 2;

      const username = String(row["UserName "] || "").trim();
      const amount = Number(row["Amount "] || 0);

      const expenseMonthRaw = String(row["Expense Month "] || "");
      const expenseMonthStr = capitalizeMonth(expenseMonthRaw);

      const expMonth = getMonthNumber(expenseMonthStr);

      let rowErrors: string[] = [];

      for (const col of requiredColumns) {
        if (col === "Remarks (if any)") {
          continue;
        }

        if (!row[col] || row[col].toString().trim() === "") {
          rowErrors.push(`${col} is required`);
        }
      }

      if (isNaN(amount)) {
        rowErrors.push("Amount must be numeric");
      }

      if (amount < 0) {
        rowErrors.push("Amount cannot be negative");
      }

      let isValidMonth = false;

      if (expMonth !== undefined) {
        if (expMonth === currentMonth) {
          isValidMonth = true;
        }

        const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;

        if (expMonth === prevMonth && currentDate <= 5) {
          isValidMonth = true;
        }

        if (expMonth > currentMonth) {
          isValidMonth = true;
        }
      }

      if (!isValidMonth) {
        rowErrors.push("Invalid Expense Month (past month not allowed)");
      }

      if (rowErrors.length > 0) {
        errorList.push({
          row: rowNumber,
          data: row,
          errors: rowErrors,
        });
      } else {
        validRows.push({
          Title: username,
          Username: username,
          EmployeeCostCenter: String(row["Cost Center "] || ""),
          EmployeeCostCenterName: String(
            row["Cost Center Name "] || "",
          ),
          VendorName: String(row["Vendor Name"] || ""),
          VendorCode: String(row["Vendor Code "] || ""),
          PONumber: String(row["PO Number"] || ""),
          GLCode: String(row["GL Code "] || ""),
          GLDescription: String(row["GL Description "] || ""),
          Amount: amount,
          ExpenseMonth: expenseMonthStr,
          Remarks: String(row["Remarks (if any)"] || ""),
          Status: "Pending",
        });
      }
    }

    const escodata = (v: any) => String(v ?? "").replace(/'/g, "''");

    const buildFullKey = (item: any) =>
      [
        item.Username,
        item.EmployeeCostCenter,
        item.EmployeeCostCenterName,
        item.VendorName,
        item.VendorCode,
        item.PONumber,
        item.GLCode,
        item.GLDescription,
        item.Amount,
        item.ExpenseMonth,
        item.Remarks,
      ]
        .map((v) => String(v ?? "").trim().toLowerCase())
        .join("||");

    try {
      const displayData: any[] = [];

      const duplicateKeys = new Set<string>();

      for (const item of validRows) {
        const key = buildFullKey(item);

        let isDuplicate = false;

        if (duplicateKeys.has(key)) {
          isDuplicate = true;
        } else {
          duplicateKeys.add(key);
        }

        const filterQuery = [
          `Username eq '${escodata(item.Username)}'`,
          `VendorCode eq '${escodata(item.VendorCode)}'`,
          `PONumber eq '${escodata(item.PONumber)}'`,
          `GLCode eq '${escodata(item.GLCode)}'`,
          `ExpenseMonth eq '${escodata(item.ExpenseMonth)}'`,
          `Amount eq ${Number(item.Amount) || 0}`,
        ].join(" and ");

        const existingItems = await sp.web.lists
          .getByTitle("AccrualSheetList")
          .items.select(
            "Username",
            "EmployeeCostCenter",
            "EmployeeCostCenterName",
            "VendorName",
            "VendorCode",
            "PONumber",
            "GLCode",
            "GLDescription",
            "Amount",
            "ExpenseMonth",
            "Remarks",
          )
          .filter(filterQuery)();

        const norm = (v: any) => String(v ?? "").trim().toLowerCase();

        const hasServerDuplicate = existingItems.some(
          (existing: any) =>
            norm(existing.EmployeeCostCenter) ===
              norm(item.EmployeeCostCenter) &&
            norm(existing.EmployeeCostCenterName) ===
              norm(item.EmployeeCostCenterName) &&
            norm(existing.VendorName) === norm(item.VendorName) &&
            norm(existing.GLDescription) === norm(item.GLDescription) &&
            norm(existing.Remarks) === norm(item.Remarks),
        );

        if (hasServerDuplicate) {
          isDuplicate = true;
        }

        await sp.web.lists.getByTitle("AccrualSheetList").items.add(item);

        displayData.push({
          ...item,
          isDuplicate,
        });
      }

      console.log("Submitted Data", displayData);

      setSubmittedData(displayData);

      setErrors(errorList);

      if (validRows.length > 0 && errorList.length > 0) {
        alert(
          `✅ ${validRows.length} records saved\n❌ ${errorList.length} records failed`,
        );
      } else if (validRows.length > 0) {
        alert("All records saved successfully ✅");
      } else {
        alert("No valid data to save ❌");
      }

      setExcelData([]);
      setFile(null);
    } catch (error) {
      console.log("Save Error", error);
      alert("Error saving records");
    } finally {
      setIsSubmitting(false);
    }
  };
  const handleExit = () => {
          const webUrl = props.context.pageContext.web.absoluteUrl;

    window.location.href = `${webUrl}/SitePages/Accuralsheet.aspx`;
  };
  const exitPage1 = async () => {
    window.location.href = `${window.location.origin}/sites/SonaFinance/SitePages/Accuralsheet.aspx`;
  };
  const Resetpage = () => {
    setExcelData([]);
    setFile(null);

    setErrors([]);
    setData([]);
    setFilteredData([]);
    setIsSearched(false);
  };

  const downloadTemplate1 = async () => {
    try {
      const files = await sp.web.lists
        .getByTitle("Accrualtemplate")
        .items.select("FileRef", "FileLeafRef", "Modified")
        .orderBy("Modified", false)
        .top(1)();

      if (files.length > 0) {
        window.open(files[0].FileRef, "_blank");
      }
    } catch (error) {
      console.log("Download error:", error);
    }
  };
  const uploadFile = async () => {
    if (!file) {
      alert("Please select file");
      return;
    }

    try {
      const fileBuffer = await file.arrayBuffer();

      const workbook = XLSX.read(fileBuffer, { type: "array" });

      const sheetName = workbook.SheetNames[0];
      const sheet = workbook.Sheets[sheetName];

      const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      const isValid = validateTemplate(data);

      if (!isValid) {
        setExcelData([]);
        alert("Invalid template. Please use the correct format.");
        return;
      }

      setExcelData(data);

      alert("File validated and uploaded successfully");
    } catch (error) {
      console.log("Upload error:", error);
    }
  };

  const uploadFile1 = async () => {
    if (!file) {
      alert("Please select file");
      return;
    }

    try {
      const fileBuffer = await file.arrayBuffer();
      const fileName = file.name;

      await sp.web
        .getFolderByServerRelativePath(
          "/sites/SonaFinance/UploadAccuralTemplate",
        )
        .files.addUsingPath(fileName, fileBuffer, { Overwrite: true });

      alert("File uploaded successfully");

      const workbook = XLSX.read(fileBuffer, { type: "array" });

      const sheetName = workbook.SheetNames[0];

      const sheet = workbook.Sheets[sheetName];

      const data = XLSX.utils.sheet_to_json(sheet, { defval: "" });

      setExcelData(data);
    } catch (error) {
      console.log("Upload error:", error);
    }
  };
  React.useEffect(() => {
    void loadDuplicateData();
  }, []);
  return (
    <div>
      <div className="header">
        <h1>Upload Accrual Sheet</h1>
      </div>
      <div style={{ padding: "10px" }}>
        <div className="heading1">
          <label>Requestor Information</label>
        </div>
        <div className="main-formcontainer">
          <div className="row mb-20">
            <div className="col-md-4">
              <label htmlFor="Employee Code" className="font">
                Employee Code
              </label>
              <input
                value={employee.EmployeeCode || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="Employee Name" className="font">
                Employee Name{" "}
              </label>
              <input
                value={employee.EmployeeName || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="Employee Email" className="font">
                Employee Email{" "}
              </label>
              <input
                value={employee.Email || ""}
                className="form-control readonly"
              />
            </div>
          </div>
          <div className="row mb-20">
            <div className="col-md-4">
              <label htmlFor="Contact No" className="font">
                Contact No
              </label>
              <input
                value={employee.ContactNo || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="Employee Status" className="font">
                Employee Status
              </label>
              <input
                value={employee.EmployeeStatus || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="Division" className="font">
                Division
              </label>
              <input
                value={employee.Division || ""}
                className="form-control readonly"
              />
            </div>
          </div>
          <div className="row mb-20">
            <div className="col-md-4">
              <label htmlFor="Location" className="font">
                Location
              </label>
              <input
                value={employee.Location || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="RM" className="font">
                RM
              </label>
              <input
                value={employee.RM || ""}
                className="form-control readonly"
              />
            </div>
            <div className="col-md-4">
              <label htmlFor="HOD" className="font">
                HOD
              </label>
              <input
                value={employee.HOD|| ""}
                className="form-control readonly"
              />
            </div>
          </div>
        </div>
      </div>
      <div style={{ padding: "0px 10px" }}>
        <div className="row">
          <div className="col-md-4">
            <div>
              <button className="primaryBtn" onClick={downloadTemplate}>
                Download Template
              </button>
              <p style={{ color: "red", fontSize: "12px" }}>
                Download the above template to upload the data
              </p>
            </div>
          </div>
        </div>
        <div className="row">
          <div
            className="col-md-12"
            style={{
              display: "flex",
              gap: "15px",
              margin: "0px 10px",
              alignItems: "center",
            }}
          >
            <div>
              <label>Select Excel File</label>
              <input type="file" accept=".xlsx" onChange={handleFileChange} />
              {file && <p>Selected: {file.name}</p>}
            </div>
            <div>
              <button className="primaryBtn" onClick={uploadFile}>
                {" "}
                Upload Accrual Sheet
              </button>
            </div>
          </div>
        </div>
      </div>
      {excelData.length > 0 && (
        <div style={{ overflowX: "auto" }}>
          <table className="Custom-table">
            <thead>
              <tr>
                <th>
                  <input
                    type="checkbox"
                    checked={
                      excelData.length > 0 &&
                      selectedRows.length === excelData.length
                    }
                    onChange={(e) => handleSelectAll(e.target.checked)}
                  />
                </th>

                {Object.keys(excelData[0]).map((key) => (
                  <th key={key}>{key}</th>
                ))}

                <th>Actions</th>
              </tr>
            </thead>

            <tbody>
              {excelData.map((row: any, index: number) => (
                <tr key={index}>
                  <td>
                    <input
                      type="checkbox"
                      checked={selectedRows.includes(index)}
                      onChange={() => handleRowSelect(index)}
                    />
                  </td>

                  {Object.keys(row).map((key) => (
                    <td key={key}>
                      {editingRow === index ? (
                        <input
                          value={row[key]}
                          onChange={(e) =>
                            handleCellChange(index, key, e.target.value)
                          }
                        />
                      ) : (
                        row[key]
                      )}
                    </td>
                  ))}

                  <td>
                    <div
                      style={{
                        display: "flex",
                        gap: "5px",
                        alignItems: "center",
                        justifyContent: "center",
                      }}
                    >
                      {editingRow === index ? (
                        <button onClick={() => saveRow()}>Save</button>
                      ) : (
                        <a onClick={() => setEditingRow(index)}>
                          <img src={edit} width={15} height={15} />
                        </a>
                      )}

                      <a
                        onClick={() => deleteRow(index)}
                        style={{ marginLeft: "5px" }}
                      >
                        <img src={del} width={15} height={15} />
                      </a>
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      {errors.length > 0 && (
        <div style={{ marginTop: "30px" }}>
          <h3 style={{ color: "red" }}>Validation Errors</h3>

          <table
            style={{
              width: "100%",
              borderCollapse: "collapse",
              marginTop: "10px",
            }}
          >
            <thead style={{ background: "#ff4d4f", color: "#fff" }}>
              <tr>
                <th>Row</th>
                <th>
                  <input
                    type="checkbox"
                    checked={
                      errors.length > 0 && selectedRows.length === errors.length
                    }
                    onChange={(e) => handleSelectAll(e.target.checked)}
                  />
                </th>
                <th>User</th>
                <th>Cost Center</th>
                <th>Cost Center Name</th>
                <th>Vendor</th>
                <th>PO</th>
                <th>Amount</th>
                <th>Month</th>
                <th>Errors</th>
              </tr>
            </thead>

            <tbody>
              {errors.map((err, index) => (
                <tr key={index}>
                  <td>
                    <input
                      type="checkbox"
                      checked={selectedRows.includes(index)}
                      onChange={() => handleRowSelect(index)}
                    />
                  </td>

                  <td>{err.row}</td>

                  <td>{err.data["UserName "]}</td>
                  <td>{err.data["Employee Cost Center "]}</td>
                  <td>{err.data["Employee Cost Center Name "]}</td>
                  <td>{err.data["Vendor Name"]}</td>
                  <td>{err.data["PO Number"]}</td>
                  <td>{err.data["Amount "]}</td>
                  <td>{err.data["Expense Month "]}</td>

                  <td style={{ color: "red" }}>{err.errors.join(", ")}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}

      {excelData.length > 0 && (
        <button
          className="sendback-btn"
          onClick={deleteSelectedRows}
          disabled={selectedRows.length === 0}
        >
          Delete Selected
        </button>
      )}

      {submittedData.length > 0 && (
        <div style={{ marginTop: "20px" }}>
          <h3>Submitted Records</h3>

          <table className="table table-bordered">
            <thead>
              <tr>
                <th>User</th>
                <th>Cost Center</th>
                <th>Cost Center Name</th>
                <th>Vendor</th>
                <th>PO Number</th>
                <th style={{width:"150px"}} >Amount</th>
                <th>Month</th>
                <th>Status</th>
              </tr>
            </thead>

            <tbody>
              {submittedData.map((row, index) => (
                <tr
                  key={index}
                  style={{
                    backgroundColor: row.isDuplicate ? "#ffe6e6" : "white",
                  }}
                >
                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.Username}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.EmployeeCostCenter}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.EmployeeCostCenterName}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.VendorName}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.PONumber}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.Amount}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "black",
                      fontWeight: row.isDuplicate ? "bold" : "normal",
                    }}
                  >
                    {row.ExpenseMonth}
                  </td>

                  <td
                    style={{
                      color: row.isDuplicate ? "red" : "green",
                      fontWeight: "bold",
                    }}
                  >
                    {row.isDuplicate ? "Duplicate Record" : "New Record"}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
      <div
        style={{
          display: "flex",
          gap: "5px",
          padding: "8px",
          justifyContent: "center",
        }}
      >
        <div>
          <button
            className="submit-btn"
            onClick={submitData}
            disabled={excelData.length === 0 || isSubmitting}
            style={{
              opacity: excelData.length === 0 || isSubmitting ? 0.5 : 1,
              cursor:
                excelData.length === 0 || isSubmitting
                  ? "not-allowed"
                  : "pointer",
            }}
          >
            {isSubmitting ? "Submitting..." : "Submit"}
          </button>
        </div>
        <div>
          <button className="reset-btn" onClick={Resetpage}>
            Reset
          </button>
        </div>
        <div>
          <button onClick={handleExit} className="Reject-btn">
            {" "}
            Exit{" "}
          </button>
        </div>
      </div>
    </div>
  );
}