import * as React from "react";
import { useState, useEffect } from "react";
import { sp } from "../DataUploadWebPart";
import * as XLSX from "xlsx";
import { saveAs } from "file-saver";
import "../components/Site.scss";
import { Link, useHistory } from "react-router-dom";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import Left from "../assets/LeftArrow.png";
import Right from "../assets/RightArrow.png";
import { IDataUploadProps } from "./IDataUploadProps";

interface IData {
  Id: number;
  Username: string;
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
}

export default function AccuralReport(props: IDataUploadProps) {
  const [isPerformer, setIsPerformer] = React.useState(false);
  const [file, setFile] = React.useState<File | null>(null);
  const [data, setData] = React.useState<IData[]>([]);
  const [isSearched, setIsSearched] = React.useState(false);
  const [userName, setUserName] = React.useState("");
  const [vendorName, setVendorName] = React.useState("");
  const [poNumber, setPoNumber] = React.useState("");
  const [fromDate, setFromDate] = React.useState("");
  const [toDate, setToDate] = React.useState("");
  const [userOptions, setUserOptions] = React.useState<string[]>([]);
  const [vendorOptions, setVendorOptions] = React.useState<string[]>([]);
  const [poOptions, setPoOptions] = React.useState<string[]>([]);
  const [selectedIds, setSelectedIds] = useState<number[]>([]);
  const [filteredData, setFilteredData] = useState<any[]>([]);

  // Pagination
  const itemsPerPage = 50;
  const [currentPage, setCurrentPage] = useState(1);
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  const history = useHistory();

  // SEARCH DATA
  const Reset = () => {
    setUserName("");
    setVendorName("");
    setPoNumber("");
    setFromDate("");
    setToDate("");

    setData([]);
    setIsSearched(false);
  };
  const loadDropdownData = async () => {
    try {
      const items: any[] = await sp.web.lists
        .getByTitle("AccrualSheetList")
        .items.select("Username", "VendorName", "PONumber")
        .top(5000)();

      // ✅ FIXED unique function (NO Set, NO spread)
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
  const handleExit = () => {
    
   // window.location.href = `${window.location.origin}/sites/SonaFinance/SitePages/Accuralsheet.aspx`;
         const webUrl =props.context.pageContext.web.absoluteUrl;
    window.location.href = `${webUrl}/SitePages/Accuralsheet.aspx`;
    //https://sonacomstargroup.sharepoint.com/sites/RLY_Finance_UAT/SitePages/Accuralsheet.aspx
    // window.location.href = `https://sonacomstargroup.sharepoint.com/sites/RLY_Finance_UAT/SitePages/Accuralsheet.aspx`;
  };
  const searchData = async () => {
    let filter = [];

    // Hide soft deleted records
    filter.push("DeleteFlag ne 1");

    if (userName) filter.push(`substringof('${userName}', Username)`);
    if (vendorName) filter.push(`substringof('${vendorName}', VendorName)`);
    if (poNumber) filter.push(`substringof('${poNumber}', PONumber)`);

    // Date filter
    // From Date
    if (fromDate) {
      filter.push(`Created ge datetime'${new Date(fromDate).toISOString()}'`);
    }

    // To Date (FIXED)
    if (toDate) {
      const to = new Date(toDate);
      to.setDate(to.getDate() + 1);

      filter.push(`Created lt datetime'${to.toISOString()}'`);
    }

    const query = filter.join(" and ");

    const items = await sp.web.lists
      .getByTitle("AccrualSheetList")
      .items.filter(query)
      .top(5000)();

    setIsSearched(true);
    setData(items);
    setFilteredData(items);
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

  // DELETE ROW
  
  const formatMonth = (value: any) => {
    if (!value) return "";

    // If it's a valid date → convert to month name
    const date = new Date(value);
    if (!isNaN(date.getTime())) {
      return date.toLocaleString("en-US", { month: "long" });
    }

    // If already string → capitalize properly
    const str = String(value).trim().toLowerCase();
    return str.charAt(0).toUpperCase() + str.slice(1);
  };

  // EXPORT CSV
  // const exportExcel = ()=>{

  // const csv = data.map(r=>Object.values(r).join(",")).join("\n")

  // const blob = new Blob([csv])
  // const url = URL.createObjectURL(blob)

  // const a = document.createElement("a")
  // a.href = url
  // a.download = "AccrualReport.csv"
  // a.click()

  // }

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };

  const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage,
  );

  React.useEffect(() => {
    void checkUserAccess();
    void loadDropdownData(); // ✅ ADD THIS
  }, []);

  
  return (
    <div>
      <div className="header">
        <h1>Accrual Report</h1>
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
                    <th className="px-4 py-2">UserName</th>
                    <th className="px-4 py-2">Cost Center</th>
                    <th className="px-4 py-2">Cost Center Name</th>
                    <th className="px-4 py-2">Vendor Name</th>
                    <th className="px-4 py-2">Vendor Code</th>
                    <th className="px-4 py-2">PO Number</th>
                    <th className="px-4 py-2">GL Code</th>
                    <th className="px-4 py-2">GL Description</th>
                    <th className="px-4 py-2">Amount</th>
                    <th className="px-4 py-2">Expense Month</th>
                    <th className="px-4 py-2">Remarks</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.map((item, index) => (
                    <tr key={index} className="border-t">
                     
                      <td className="px-4 py-2">{item.Username}</td>
                      <td className="px-4 py-2">{item.EmployeeCostCenter}</td>
                      <td className="px-4 py-2">{item.EmployeeCostCenterName}</td>
                      <td className="px-4 py-2">{item.VendorName}</td>
                      <td className="px-4 py-2">{item.VendorCode}</td>
                      <td className="px-4 py-2">{item.PONumber}</td>
                      <td className="px-4 py-2">{item.GLCode}</td>
                      <td className="px-4 py-2">{item.GLDescription}</td>
                      <td className="px-4 py-2">{item.Amount}</td>
                      <td className="px-4 py-2">
                        {formatMonth(item.ExpenseMonth)}
                      </td>

                      {/* <td className="px-4 py-2">{item.ExpenseMonth}</td> */}
                      <td className="px-4 py-2">{item.Remarks}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            <div className="flex justify-center mt-6 overflow-x-auto">
              <div
                className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow"
                style={{ textAlign: "end" }}
              >
                {/* Previous Button */}
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
                {/* Main Page Numbers */}
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

                {/* Next Button */}
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
          <div style={{ textAlign: "center", marginTop: "10px" }}>
            <button onClick={handleExit} className="Reject-btn">
              {" "}
              Exit{" "}
            </button>
          </div>
        )}
      </div>
    </div>
  );
}