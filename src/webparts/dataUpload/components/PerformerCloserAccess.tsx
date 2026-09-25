import * as React from "react";
import { sp } from "../DataUploadWebPart";

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import { useState, useEffect } from "react";
import Left from "../assets/LeftArrow.png";
import Right from "../assets/RightArrow.png";
import { IDataUploadProps } from "./IDataUploadProps";

interface IData {
  FinancialYear: string;
  Month: string;
  DateofClosure: string;
  CreatedBy: string;
  CreatedOn: string;
}

export default function PerformerClosureAccess(props: IDataUploadProps) {

  const [division, setDivision] = React.useState("");
  const [fromDate, setFromDate] = React.useState("");
  const [uptoDate, setUptoDate] = React.useState("");
  const [historyData, setHistoryData] = React.useState<IData[]>([]);
const [isSaving, setIsSaving] = React.useState(false);

  const [filteredData, setFilteredData] = useState<any[]>([]);

  // Pagination
  const itemsPerPage = 50;
  const [currentPage, setCurrentPage] = useState(1);
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  // ✅ Financial Year dropdown
  const getFinancialYears = () => {
    const currentYear = new Date().getFullYear();
    const years = [];

    for (let i = 0; i < 5; i++) {
      const start = currentYear - i;
      const end = start + 1;
      years.push(`${start}-${end}`);
    }

    return years;
  };
  const handleExit = () => {
    //https://isriglobal.sharepoint.com/sites/SonaFinance/_layouts/workbench.aspx
  // window.location.href = `${window.location.origin}/sites/SonaFinance/SitePages/Accuralsheet.aspx`;
        const webUrl = props.context.pageContext.web.absoluteUrl;
    window.location.href = `${webUrl}/SitePages/Accuralsheet.aspx`;
  };
  // ✅ Fetch data
  const getHistoryData = async () => {
    try {
      const items = await sp.web.lists
        .getByTitle("PerformerClosureAccess")
        .items
        .select("FinancialYear", "Month", "DateofClosure", "Created", "Author/Title")
        .expand("Author")
        .orderBy("Created", false)
        .top(5000)();

      const data = items.map((item: any) => ({
        FinancialYear: item.FinancialYear,
        Month: item.Month,
        DateofClosure: new Date(item.DateofClosure).toLocaleDateString(),
        CreatedBy: item.Author?.Title,
        CreatedOn: new Date(item.Created).toLocaleDateString()
      }));

      setHistoryData(data);
      setFilteredData(data);

    } catch (error) {
      console.log("Fetch error:", error);
    }
  };
const handleSave = async () => {

  if (isSaving) return; // 🚫 prevent double click

  if (!division || !fromDate || !uptoDate) {
    alert("All fields are required");
    return;
  }

  try {
    setIsSaving(true); // ✅ disable button

    await sp.web.lists
      .getByTitle("PerformerClosureAccess")
      .items.add({
        Title: division + " - " + fromDate,
        FinancialYear: division,
        Month: fromDate,
        DateofClosure: uptoDate
      });

    alert("Saved successfully ✅");

    // reset
    setDivision("");
    setFromDate("");
    setUptoDate("");

    await getHistoryData();

  } catch (error) {
    console.log(error);
    alert("Error saving data");
  } finally {
    setIsSaving(false); // ✅ enable again
  }
};

  const th: React.CSSProperties = {
    border: "1px solid #ccc",
    padding: "8px",
    textAlign: "center"
  };

  const td: React.CSSProperties = {
    border: "1px solid #ccc",
    padding: "8px",
    textAlign: "center"
  };

  const handlePageChange = (page: number) => {
    if (page >= 1 && page <= totalPages) {
      setCurrentPage(page);
    }
  };



  const sortedData = [...filteredData].sort((a, b) => b.ID - a.ID);

  const paginatedData = sortedData.slice(
    (currentPage - 1) * itemsPerPage,
    currentPage * itemsPerPage
  );

  // ✅ Load on page load
  React.useEffect(() => {
    void getHistoryData();
  }, []);

  // ✅ Save function
  

  return (
    <div>
      <div className='header'>
        <h1>Performer Closure Access</h1>
      </div>
      <div className="PaddAll">

        <div className='row mb-20'>
          <div className='col-md-4'>
            <label htmlFor="Financial Year" className='font'>Financial Year</label>
            <select value={division} className="form-control" onChange={(e) => setDivision(e.target.value)}>
              <option value="">Select Year</option>
              {getFinancialYears().map((yr) => (
                <option key={yr} value={yr}>{yr}</option>
              ))}
            </select>
          </div>
          <div className='col-md-4'>
            <label htmlFor="Month" className='font'>Month</label>
            <select value={fromDate} className="form-control" onChange={(e) => setFromDate(e.target.value)}>
              <option value="">Select Month</option>
              {["january", "february", "march", "april", "may", "june", "july", "august", "september", "october", "november", "december"]
                .map(m => <option key={m} value={m}>{m}</option>)}
            </select>
          </div>
          <div className='col-md-4'>
            <label htmlFor="Employee Email" className='font'>Date of Closing</label>
            <input type="date" className="form-control" value={uptoDate} onChange={e => setUptoDate(e.target.value)} />
          </div>
        </div>
        <div className="row mb-20">
          <div className="col-md-12" style={{ margin: "10px", textAlign: "end" }}>
          <button
  onClick={handleSave}
  disabled={isSaving}
  style={{
    opacity: isSaving ? 0.5 : 1,
    cursor: isSaving ? "not-allowed" : "pointer"
  }}
>
  {isSaving ? "Saving..." : "Save"}
</button>

          </div>
        </div>

        {/* TABLE */}
        <hr style={{ marginTop: "30px" }} />


        <div className="heading1">
          <label>History Info</label>
        </div>
        <div className='main-formcontainer'>
          <div className="overflow-x-auto">
            <div className="table-vert-scroll">

              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  style={{ backgroundColor: "#3c3e45" }}
                  className="text-white"
                >
                  <tr>
                    <th className="px-4 py-2">S.No</th>
                    <th className="px-4 py-2">Financial Year</th>
                    <th className="px-4 py-2">Month</th>
                    <th className="px-4 py-2">Closing Date</th>
                    <th className="px-4 py-2">Created By</th>
                    <th style={th}>Created On</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.map((item, index) => (
                    <tr key={index} className="border-t">
                      <td className="px-4 py-2">{index + 1}</td>
                      <td className="px-4 py-2">{item.FinancialYear}</td>
                      <td className="px-4 py-2">{item.Month}</td>
                      <td className="px-4 py-2">{item.DateofClosure}</td>
                      <td className="px-4 py-2">{item.CreatedBy}</td>
                      <td className="px-4 py-2">{item.CreatedOn}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            {/* Pagination */}
            <div className="flex justify-center mt-6 overflow-x-auto">
              <div className="flex space-x-2 flex-nowrap px-4 py-2 bg-#2149d5 rounded shadow" style={{ textAlign: "end" }}>
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
                        backgroundColor: currentPage === page ? "#3c3e45" : "#fff",
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
        </div>

        <div style={{ textAlign: "center", marginTop: "10px" }}>
          <button onClick={handleExit} className="Reject-btn"> Exit </button>
        </div>
      </div>

    </div>
  );
}
