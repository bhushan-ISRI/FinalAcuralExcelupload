import * as React from "react";
import { sp } from "../DataUploadWebPart";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import {
  TextField,
  DefaultButton,
  Dropdown,
  IDropdownOption,
} from "@fluentui/react";

import { useState, useEffect } from "react";
import Left from "../assets/LeftArrow.png";
import Right from "../assets/RightArrow.png";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import { IDataUploadProps } from "./IDataUploadProps";

interface IData {
  Id: number;
  Username: string;
  Created: string;
  VendorName: string;
  VendorCode: string;
  PONumber: string;
  GLCode: string;
  GLDescription: string;
  EmployeeCostCenter: string;
  EmployeeCostCenterName: string;
  Amount: number;
  ExpenseMonth: string;
  Remarks: string;
  Status?: string;
}

const toUtcMidnight = (dateStr: string): Date =>
  new Date(`${dateStr}T00:00:00.000Z`);

export default function AdjustmentReport(props: IDataUploadProps) {
  const [isPerformer, setIsPerformer] = React.useState(false);
  const [file, setFile] = React.useState<File | null>(null);
  const [data, setData] = React.useState<IData[]>([]);
  const [isSearched, setIsSearched] = React.useState(false);
  const [closedMonths, setClosedMonths] = React.useState<string[]>([]);
  const [editingIds, setEditingIds] = useState<number[]>([]);
  const [userName, setUserName] = React.useState("");
  const [vendorName, setVendorName] = React.useState("");
  const [poNumber, setPoNumber] = React.useState("");
  const [fromDate, setFromDate] = React.useState("");
  const [toDate, setToDate] = React.useState("");
  const [userOptions, setUserOptions] = React.useState<string[]>([]);
  const [vendorOptions, setVendorOptions] = React.useState<string[]>([]);
  const [poOptions, setPoOptions] = React.useState<string[]>([]);

  const [filteredData, setFilteredData] = useState<any[]>([]);

  const itemsPerPage = 10;
  const [currentPage, setCurrentPage] = useState(1);
  const totalPages = Math.ceil(data.length / itemsPerPage);
  const [selectedIds, setSelectedIds] = useState<number[]>([]);

  const handleExit = () => {
          const webUrl = props.context.pageContext.web.absoluteUrl;

    window.location.href = `${webUrl}/SitePages/Accuralsheet.aspx`;
  };

  const searchData = async () => {
    let filter: string[] = [];

    console.log("🔍 [searchData] Raw filter inputs:", {
      userName: JSON.stringify(userName),
      vendorName: JSON.stringify(vendorName),
      poNumber: JSON.stringify(poNumber),
      fromDate: JSON.stringify(fromDate),
      toDate: JSON.stringify(toDate),
    });

    filter.push("DeleteFlag ne 1");

    filter.push("(Status eq 'Pending' or Status eq 'pending')");

    if (userName) filter.push(`substringof('${userName}', Username)`);
    if (vendorName) filter.push(`substringof('${vendorName}', VendorName)`);
    if (poNumber) filter.push(`substringof('${poNumber}', PONumber)`);

    if (fromDate) {
      const from = toUtcMidnight(fromDate);
      filter.push(`Created ge datetime'${from.toISOString()}'`);
    }

    if (toDate) {
      const to = toUtcMidnight(toDate);
      to.setUTCDate(to.getUTCDate() + 1);
      filter.push(`Created lt datetime'${to.toISOString()}'`);
    }

    const query = filter.join(" and ");

    console.log("🔍 [searchData] Filter clauses:", filter);
    console.log("🔍 [searchData] Final OData $filter query:", query);

    try {
      const items = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.filter(query)
        .top(5000)();

      console.log("🔍 [searchData] Rows returned by REST call:", items.length);
      console.log("All Items:", items);

      if (items.length > 0) {
        console.log(
          "🔍 [searchData] Sample field values from first 5 rows:",
          items.slice(0, 5).map((i: any) => ({
            Username: JSON.stringify(i.Username),
            VendorName: JSON.stringify(i.VendorName),
            PONumber: JSON.stringify(i.PONumber),
            Status: JSON.stringify(i.Status),
            ExpenseMonth: JSON.stringify(i.ExpenseMonth),
            Created: i.Created,
          })),
        );
      } else {
        console.log(
          "🔍 [searchData] REST call returned ZERO rows — the $filter query itself excluded everything. " +
            "Check the 'Final OData $filter query' log above.",
        );
      }

      const formatMonth = (value: any) => {
        if (!value) return "";
        const str = String(value).trim().toLowerCase();
        return str.charAt(0).toUpperCase() + str.slice(1);
      };

      const filteredData = items
        .filter((item: any) => {
          const status = String(item.Status || "")
            .trim()
            .toLowerCase();

          if (status !== "pending") {
            console.log(
              `🔍 [client filter] DROPPED row Id=${item.Id} (${item.Username}) — status was "${status}", expected "pending"`,
            );
            return false;
          }

          return true;
        })
        .map((item: any) => ({
          ...item,
          ExpenseMonth: formatMonth(item.ExpenseMonth),
        }));

      console.log(
        "🔍 [searchData] Rows AFTER client-side status filter (no month-closure check):",
        filteredData.length,
      );
      console.log("Final Data:", filteredData);

      setIsSearched(true);
      setData(filteredData);
      setFilteredData(filteredData);
    } catch (error) {
      console.error("Search Error:", error);
      alert("Error fetching data");
    }
  };

  const removeNewRow = (index: number) => {
    const updated = [...data];
    updated.splice(index, 1);
    setData(updated);
  };
  const getClosureData = async () => {
    const items = await sp.web.lists
      .getByTitle("PerformerClosureAccess")
      .items.select("Month", "DateofClosure")
      .top(5000)();

    return items.map((item: any) => ({
      month: item.Month?.toLowerCase(),
      closingDate: new Date(item.DateofClosure),
    }));
  };

  const deleteSelectedItems = async () => {
    if (selectedIds.length === 0) {
      alert("Please select records");
      return;
    }

    const confirmDelete = window.confirm(
      `Are you sure you want to delete ${selectedIds.length} record(s)?`,
    );

    if (!confirmDelete) return;

    try {
      for (const id of selectedIds) {
        await sp.web.lists
          .getByTitle("AccrualSheetList")
          .items.getById(id)
          .update({
            DeleteFlag: true,
          });
      }

      alert("Selected records deleted successfully");

      setSelectedIds([]);

      await searchData();
    } catch (error) {
      console.log(error);
      alert("Error deleting records");
    }
  };

  const updateGLCode = async (item: IData) => {
    try {
      await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.getById(item.Id)
        .update({
          GLCode: item.GLCode,
        });

      alert("GL Code updated successfully");

      setEditingIds(editingIds.filter((id) => id !== item.Id));

      await searchData();
    } catch (error) {
      console.log(error);
      alert("Error updating GL Code");
    }
  };
  const getClosedMonths = async () => {
    const items = await sp.web.lists
      .getByTitle("PerformerClosureAccess")
      .items.select("Month", "DateofClosure")
      .top(5000)();

    return items.map((i: any) => {
      const month = String(i.Month || "")
        .trim()
        .toLowerCase();
      const year = i.DateofClosure
        ? new Date(i.DateofClosure).getFullYear()
        : "";

      return `${month}-${year}`;
    });
  };

  const searchData1 = async () => {
    let filter: string[] = [];

    filter.push("DeleteFlag ne 1");

    filter.push("(Status eq 'Pending' or Status eq 'pending')");

    if (userName && userName.trim() !== "") {
      filter.push(`substringof('${userName.trim()}', Username)`);
    }

    if (vendorName && vendorName.trim() !== "") {
      filter.push(`substringof('${vendorName.trim()}', VendorName)`);
    }

    if (poNumber && poNumber.trim() !== "") {
      filter.push(`substringof('${poNumber.trim()}', PONumber)`);
    }

    if (fromDate) {
      const from = toUtcMidnight(fromDate);
      filter.push(`Created ge datetime'${from.toISOString()}'`);
    }

    if (toDate) {
      const to = toUtcMidnight(toDate);
      to.setUTCDate(to.getUTCDate() + 1);
      filter.push(`Created lt datetime'${to.toISOString()}'`);
    }

    const query = filter.join(" and ");

    try {
      const items = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.filter(query)
        .top(5000)();

      console.log("All Items:", items);

      const closedItems = await sp.web.lists
        .getByTitle("PerformerClosureAccess")
        .items.select("Month")
        .top(5000)();

      const closedMonths = closedItems.map((i: any) =>
        String(i.Month || "")
          .trim()
          .toLowerCase(),
      );

      console.log("Closed Months:", closedMonths);

      const formatMonth = (value: any) => {
        if (!value) return "";

        const str = String(value).trim().toLowerCase();
        return str.charAt(0).toUpperCase() + str.slice(1);
      };

      const filteredData = items
        .filter((item: any) => {
          const month = String(item.ExpenseMonth || "")
            .trim()
            .toLowerCase();

          const status = String(item.Status || "")
            .trim()
            .toLowerCase();

          console.log("Checking → Month:", month, "Status:", status);

          if (status !== "pending") return false;

          return !closedMonths.includes(month);
        })
        .map((item: any) => ({
          ...item,
          ExpenseMonth: formatMonth(item.ExpenseMonth),
        }));

      console.log("Final Data:", filteredData);

      setIsSearched(true);
      setData(filteredData);
    } catch (error) {
      console.error("Search Error:", error);
      alert("Error fetching data");
    }
  };

  const formatMonth = (value: any) => {
    if (!value) return "";

    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleString("en-US", { month: "long" });
    }

    const str = String(value).trim().toLowerCase();
    return str.charAt(0).toUpperCase() + str.slice(1);
  };

  const saveAllRows = async () => {
    const newRows = data.filter((item) => item.Id === 0);

    if (newRows.length === 0) {
      alert("No new rows to save");
      return;
    }

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
      return months[monthName.toLowerCase()];
    };

    const capitalizeMonth = (month: string) => {
      const m = month.trim().toLowerCase();
      return m.charAt(0).toUpperCase() + m.slice(1);
    };

    const today = new Date();
    const currentMonth = today.getMonth();
    const currentDate = today.getDate();

    let errorList: string[] = [];
    let errorRows: any[] = [];

    for (const item of newRows) {
      const username = item.Username || "Unknown";

      if (
        !item.Username ||
        !item.VendorName ||
        !item.VendorCode ||
        !item.PONumber ||
        !item.GLCode ||
        !item.GLDescription ||
        !item.EmployeeCostCenter ||
        !item.EmployeeCostCenterName ||
        !item.Amount ||
        !item.ExpenseMonth
      ) {
        errorList.push(`${username} → Missing required fields`);
        errorRows.push(item);
        continue;
      }

      if (isNaN(Number(item.Amount))) {
        errorList.push(`${username} → Amount must be numeric`);
        errorRows.push(item);
        continue;
      }

      if (Number(item.Amount) < 0) {
        errorList.push(`${username} → Negative amount not allowed`);
        errorRows.push(item);
        continue;
      }

      const expenseMonthStr = String(item.ExpenseMonth || "")
        .trim()
        .toLowerCase();

      const expMonth = getMonthNumber(expenseMonthStr);

      if (expMonth === undefined) {
        errorList.push(`${username} → Invalid month`);
        errorRows.push(item);
        continue;
      }

      let isValid = false;

      if (expMonth === currentMonth) {
        isValid = true;
      }

      const prevMonth = currentMonth === 0 ? 11 : currentMonth - 1;

      if (expMonth === prevMonth && currentDate <= 5) {
        isValid = true;
      }

      if (expMonth > currentMonth) {
        isValid = true;
      }

      if (!isValid) {
        errorList.push(
          `${username}-${expenseMonthStr} → Past month not allowed`,
        );
        errorRows.push(item);
      }
    }

    if (errorList.length > 0) {
      alert(
        "❌ Some records are invalid. Nothing saved.\n\n" +
          errorList.join("\n"),
      );

      setData(errorRows);
      setFilteredData(errorRows);
      setIsSearched(true);

      return;
    }

    try {
      for (const item of newRows) {
        const expenseMonthStr = capitalizeMonth(item.ExpenseMonth);

        await sp.web.lists.getByTitle("AccrualSheetList").items.add({
          Title: String(item.Username || ""),

          Username: String(item.Username || ""),
          VendorName: String(item.VendorName || ""),
          VendorCode: String(item.VendorCode || ""),
          PONumber: String(item.PONumber || ""),
          GLCode: String(item.GLCode || ""),
          GLDescription: String(item.GLDescription || ""),
          EmployeeCostCenter: String(item.EmployeeCostCenter || ""),
          EmployeeCostCenterName: String(item.EmployeeCostCenterName || ""),
          Amount: Number(item.Amount || 0),

          ExpenseMonth: expenseMonthStr,

          Remarks: String(item.Remarks || ""),
          DeleteFlag: false,
          Status: "Pending",
        });
      }

      alert("All records saved successfully ✅");

      setData([]);
      setFilteredData([]);
      setIsSearched(false);

      void searchData();
    } catch (error) {
      console.log("Bulk save error:", error);
      alert("Error saving records");
    }
  };

  const loadDropdownData = async () => {
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.select("Username", "VendorName", "PONumber")
        .filter(
          "DeleteFlag ne 1 and (Status eq 'Pending' or Status eq 'pending')",
        )
        .top(5000)();

      const unique = (arr: any[], key: string) => {
        const result: string[] = [];

        arr.forEach((item) => {
          const value = item[key];

          if (value && result.indexOf(value) === -1) {
            result.push(value);
          }
        });

        return result;
      };

      setUserOptions(unique(items, "Username"));
      setVendorOptions(unique(items, "VendorName"));
      setPoOptions(unique(items, "PONumber"));
    } catch (error) {
      console.log("Dropdown load error:", error);
    }
  };

  const freezeData = async () => {
    if (data.length === 0) {
      alert("No records to freeze");
      return;
    }

    try {
      const pendingItems = data.filter((item) => item.Status === "Pending");

      if (pendingItems.length === 0) {
        alert("No Pending records found");
        return;
      }

      for (const item of pendingItems) {
        await sp.web.lists
          .getByTitle("AccrualSheetList")
          .items.getById(item.Id)
          .update({
            Status: "Freez",
            MailSent:"No"
          });
      }

    // const flowResponse = await fetch("YOUR_FLOW_URL", {
    //   method: "POST",
    //   headers: {
    //     "Content-Type": "application/json",
    //   },
    //   body: JSON.stringify({
    //     action: "FreezeAccrual",
    //     recordCount: pendingItems.length,
    //   }),
    // });

    // if (!flowResponse.ok) {
    //   throw new Error(
    //     `Power Automate failed: ${flowResponse.status}`
    //   );
    // }


      alert("Records freeze successfully ✅");

      void searchData();
    } catch (error) {
      console.log("Freeze error:", error);
      alert("Error freezing records");
    }
  };

  const updateRow = (index: number, field: string, value: any) => {
    const updated = [...data];

    updated[index] = {
      ...updated[index],
      [field]: value,
    };

    setData(updated);
  };
  const checkUserAccess = async () => {
    try {
      const user = await sp.web.currentUser();

      const groups = await sp.web.siteUsers.getById(user.Id).groups();

      const performer = groups.some((g: any) => g.Title === "AccrualPerformer");

      setIsPerformer(performer);
    } catch (error) {
      console.log("Access check error", error);
    }
  };

  const exportExcel = () => {
    const exportData = data.map((item: any) => ({
      UserName: item.Username,
      VendorName: item.VendorName,
      VendorCode: item.VendorCode,
      PONumber: item.PONumber,
      GLCode: item.GLCode,
      GLDescription: item.GLDescription,
      EmployeeCostCenter: item.EmployeeCostCenter,
      EmployeeCostCenterName: item.EmployeeCostCenterName,
      Amount: item.Amount,
      ExpenseMonth: item.ExpenseMonth,
      Remarks: item.Remarks,
    }));

    const worksheet = XLSX.utils.json_to_sheet(exportData);

    const workbook = {
      Sheets: { "Accrual Report": worksheet },
      SheetNames: ["Accrual Report"],
    };

    const excelBuffer = XLSX.write(workbook, {
      bookType: "xlsx",
      type: "array",
    });

    const blob = new Blob([excelBuffer], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    });

    saveAs(blob, "Accrual_Report.xlsx");
  };

  const deleteItem = async (id: number) => {
    const confirmDelete = window.confirm(
      "Are you sure you want to delete the data?",
    );

    if (!confirmDelete) {
      return;
    }

    try {
      await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.getById(id)
        .update({
          DeleteFlag: true,
        });

      alert("Data deleted successfully");

      void searchData();
    } catch (error) {
      console.log("Delete error:", error);
      alert("Error deleting data");
    }
  };

  const addNewRow = () => {
    const newRow: IData = {
      Id: 0,
      Username: "",
      Created: "",
      VendorName: "",
      VendorCode: "",
      PONumber: "",
      GLCode: "",
      GLDescription: "",
      EmployeeCostCenter: "",
      EmployeeCostCenterName: "",
      Amount: 0,
      ExpenseMonth: "",
      Remarks: "",
    };

    const updated = [...data, newRow];

    setData(updated);

    setCurrentPage(Math.ceil(updated.length / itemsPerPage));
  };

  const saveRow = async (item: IData) => {
    if (
      !item.Username ||
      !item.VendorName ||
      !item.VendorCode ||
      !item.PONumber ||
      !item.GLCode ||
      !item.GLDescription ||
      !item.EmployeeCostCenter ||
      !item.EmployeeCostCenterName ||
      !item.Amount ||
      !item.ExpenseMonth
    ) {
      alert("All fields are mandatory");
      return;
    }

    if (isNaN(item.Amount)) {
      alert("Amount must be numeric");
      return;
    }

    await sp.web.lists.getByTitle("AccrualSheetList").items.add({
      Username: item.Username,
      VendorName: item.VendorName,
      VendorCode: item.VendorCode,
      PONumber: item.PONumber,
      GLCode: item.GLCode,
      GLDescription: item.GLDescription,
      EmployeeCostCenter: item.EmployeeCostCenter,
      EmployeeCostCenterName: item.EmployeeCostCenterName,
      Amount: item.Amount,
      ExpenseMonth: item.ExpenseMonth,
      Remarks: item.Remarks,
      DeleteFlag: false,
      Status: "Pending",
    });

    alert("Record Saved Successfully");

    void searchData();
  };

  const Reset = () => {
    setUserName("");
    setVendorName("");
    setPoNumber("");
    setFromDate("");
    setToDate("");

    setData([]);
    setIsSearched(false);
  };

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const paginatedData = data.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage,
  );

  React.useEffect(() => {
    void checkUserAccess();
    void loadDropdownData();
  }, []);

  return (
    <div>
      <div className="header">
        <h1>Adjustment Report</h1>
      </div>
      <div className="PaddAll">
        <div style={{ display: "flex", gap: "10px", marginBottom: "20px" }}>
          <select
            value={userName}
            className="form-control"
            onChange={(e) => setUserName(e.target.value)}
          >
            <option value="">All Users</option>
            {userOptions.map((u, i) => (
              <option key={i} value={u}>
                {u}
              </option>
            ))}
          </select>

          <select
            value={vendorName}
            className="form-control"
            onChange={(e) => setVendorName(e.target.value)}
          >
            <option value="">All Vendors</option>
            {vendorOptions.map((v, i) => (
              <option key={i} value={v}>
                {v}
              </option>
            ))}
          </select>

          <select
            value={poNumber}
            className="form-control"
            onChange={(e) => setPoNumber(e.target.value)}
          >
            <option value="">All PO</option>
            {poOptions.map((p, i) => (
              <option key={i} value={p}>
                {p}
              </option>
            ))}
          </select>

          <input
            type="date"
            className="form-control"
            value={fromDate}
            onChange={(e) => setFromDate(e.target.value)}
          />
          <input
            type="date"
            className="form-control"
            value={toDate}
            onChange={(e) => setToDate(e.target.value)}
          />
        </div>
        <div
          className=""
          style={{
            display: "block",
            marginBottom: "20px",
            textAlign: "center",
          }}
        >
          <button onClick={searchData} className="submit-btn">
            Search
          </button>
          <button onClick={exportExcel} className="sendback-btn">
            Export
          </button>
          <button onClick={Reset} className="reset-btn">
            Reset
          </button>
        </div>

        {isSearched && (
          <div className="overflow-x-auto">
            <div className="table-vert-scroll">
              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  style={{ backgroundColor: "#3c3e45" }}
                  className="text-white"
                >
                  <tr>
                    <th>
                      <input
                        type="checkbox"
                        checked={
                          paginatedData.length > 0 &&
                          paginatedData.every((x) => selectedIds.includes(x.Id))
                        }
                        onChange={(e) => {
                          if (e.target.checked) {
                            const ids = paginatedData.map((x) => x.Id);
                            setSelectedIds(ids);
                          } else {
                            setSelectedIds([]);
                          }
                        }}
                      />
                    </th>
                    <th className="px-4 py-2">Created Date</th>
                    <th className="px-4 py-2">UserName</th>
                    <th className="px-4 py-2">Employee Cost Center</th>
                    <th className="px-4 py-2">Employee Cost Center Name</th>
                    <th className="px-4 py-2">Vendor Name</th>
                    <th className="px-4 py-2">Vendor Code</th>
                    <th className="px-4 py-2">PO Number</th>
                    <th className="px-4 py-2">GL Code</th>
                    <th className="px-4 py-2">GL Description</th>
                    <th className="px-4 py-2">Amount</th>
                    <th className="px-4 py-2">Expense Month</th>
                    <th className="px-4 py-2">Remarks</th>
                    {isPerformer && <th>Delete</th>}
                    {isPerformer && <th>Action</th>}
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.length === 0 ? (
                    <tr>
                      <td colSpan={13} style={{ textAlign: "center" }}>
                        No Records Found
                      </td>
                    </tr>
                  ) : (
                    <>
                      {paginatedData.map((item, index) => {
                        const actualIndex =
                          (currentPage - 1) * itemsPerPage + index;

                        return (
                          <tr key={item.Id || actualIndex}>
                            <td>
                              {item.Id !== 0 && (
                                <input
                                  type="checkbox"
                                  checked={selectedIds.includes(item.Id)}
                                  onChange={(e) => {
                                    if (e.target.checked) {
                                      setSelectedIds([...selectedIds, item.Id]);
                                    } else {
                                      setSelectedIds(
                                        selectedIds.filter(
                                          (id) => id !== item.Id,
                                        ),
                                      );
                                    }
                                  }}
                                />
                              )}
                            </td>

                            <td>
                              {item.Created
                                ? new Date(
                                    item.Created.toString(),
                                  ).toLocaleDateString("en-GB")
                                : ""}
                            </td>
                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.Username}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "Username",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.Username
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.EmployeeCostCenter}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "EmployeeCostCenter",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.EmployeeCostCenter
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.EmployeeCostCenterName}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "EmployeeCostCenterName",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.EmployeeCostCenterName
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.VendorName}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "VendorName",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.VendorName
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.VendorCode}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "VendorCode",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.VendorCode
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.PONumber}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "PONumber",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.PONumber
                              )}
                            </td>

                            <td>
                              <input
                                value={item.GLCode || ""}
                                onChange={(e) =>
                                  updateRow(
                                    actualIndex,
                                    "GLCode",
                                    e.target.value,
                                  )
                                }
                                disabled={
                                  item.Id !== 0 && !editingIds.includes(item.Id)
                                }
                              />
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.GLDescription}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "GLDescription",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.GLDescription
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  type="number"
                                  value={item.Amount}
                                  min="0"
                                  onChange={(e) => {
                                    const val = Number(e.target.value);

                                    if (val < 0) return;

                                    updateRow(actualIndex, "Amount", val);
                                  }}
                                />
                              ) : (
                                <input
                                  type="number"
                                  value={item.Amount}
                                  readOnly
                                />
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.ExpenseMonth}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "ExpenseMonth",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                formatMonth(item.ExpenseMonth)
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <input
                                  value={item.Remarks}
                                  onChange={(e) =>
                                    updateRow(
                                      actualIndex,
                                      "Remarks",
                                      e.target.value,
                                    )
                                  }
                                />
                              ) : (
                                item.Remarks
                              )}
                            </td>

                            <td>
                              {item.Id === 0 ? (
                                <button
                                  style={{ color: "orange" }}
                                  onClick={() => removeNewRow(actualIndex)}
                                >
                                  Remove
                                </button>
                              ) : (
                                isPerformer && (
                                  <button
                                    style={{ color: "red" }}
                                    onClick={() => deleteItem(item.Id)}
                                  >
                                    Delete
                                  </button>
                                )
                              )}
                            </td>
                            {isPerformer && (
                              <td>
                                {item.Id !== 0 &&
                                  (!editingIds.includes(item.Id) ? (
                                    <button
                                      onClick={() =>
                                        setEditingIds([...editingIds, item.Id])
                                      }
                                    >
                                      Edit
                                    </button>
                                  ) : (
                                    <button onClick={() => updateGLCode(item)}>
                                      Save
                                    </button>
                                  ))}
                              </td>
                            )}
                          </tr>
                        );
                      })}

                      <tr>
                        <td colSpan={11}></td>
                        {isPerformer && (
                          <>
                            <td>
                              <button
                                className="submit-btn"
                                onClick={addNewRow}
                              >
                                Add New
                              </button>
                            </td>
                            <td>
                              <button onClick={saveAllRows}>Save All</button>
                            </td>
                          </>
                        )}
                      </tr>
                    </>
                  )}
                </tbody>
              </table>
            </div>
            {isPerformer && (
              <button
                onClick={deleteSelectedItems}
                className="sendback-btn"
                style={{ backgroundColor: "red" }}
              >
                Delete Selected
              </button>
            )}
            <div className="flex justify-center mt-6 overflow-x-auto">
              <div
                className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow"
                style={{ textAlign: "end" }}
              >
                <button
                  onClick={() => handlePageChange(currentPage - 1)}
                  disabled={currentPage === 1}
                  style={{
                    backgroundColor: "#fff",
                    border: "1px solid #000 !important",
                    marginRight: "5px",
                    opacity: currentPage === 1 ? 0.5 : 1,
                  }}
                  className="px-3 py-1 border rounded"
                >
                  <img src={Left} alt="" width={15} />
                </button>
                {Array.from({ length: totalPages }, (_, i) => i + 1)
                  .filter((page) => Math.abs(page - currentPage) <= 2)
                  .map((page) => (
                    <button
                      key={page}
                      onClick={() => handlePageChange(page)}
                      style={{
                        backgroundColor:
                          currentPage === page ? "#3c3e45" : "#fff",
                        color: currentPage === page ? "#fff" : "#000",
                        fontWeight: currentPage === page ? "bold" : "normal",
                        margin: currentPage === page ? "5px" : "5px",
                      }}
                      className="px-3 py-1 border rounded"
                    >
                      {page}
                    </button>
                  ))}

                <button
                  onClick={() => handlePageChange(currentPage + 1)}
                  disabled={currentPage === totalPages}
                  style={{
                    backgroundColor: "#fff",
                    border: "1px solid #000 !important",
                    marginLeft: "5px",
                    opacity: currentPage === totalPages ? 0.5 : 1,
                  }}
                  className="px-3 py-1 border rounded"
                >
                  <img src={Right} alt="" width={15} />
                </button>
              </div>
            </div>
          </div>
        )}
        {isPerformer && (
          <div>
            <div className="row">
              <div
                className="col-md-12 col-sm-12"
                style={{
                  display: "flex",
                  gap: "15px",
                  margin: "0px 10px",
                  alignItems: "center",
                }}
              >
                <div>
                  <button
                    className="primaryBtn"
                    style={{ margin: "0px" }}
                    onClick={freezeData}
                  >
                    Freeze send to GL
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
          </div>
        )}
      </div>
    </div>
  );
}