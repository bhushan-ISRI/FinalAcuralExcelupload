import * as React from "react";
import { sp } from "../DataUploadWebPart";
import { useState, useEffect } from "react";
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/site-groups/web";
import "@pnp/sp/site-groups";
import "@pnp/sp/site-users/web";

import {
  PeoplePicker,
  PrincipalType,
} from "@pnp/spfx-controls-react/lib/PeoplePicker";
import { IPeoplePickerContext } from "@pnp/spfx-controls-react/lib/PeoplePicker";

import { IDataUploadProps } from "./IDataUploadProps";

import Left from "../assets/LeftArrow.png";
import Right from "../assets/RightArrow.png";

interface IAccessInfo {
  Division: string;
  EmployeeName: string;
  FromDate: string;
  UptoDate: string;
  CreatedBy: string;
  CreatedOn: string;
  UpdatedBy: string;
  UpdatedOn: string;
}

export default function ManageAccess(props: IDataUploadProps) {
  const [pickerKey, setPickerKey] = React.useState<number>(0);
  const [division, setDivision] = React.useState("");
  const [fromDate, setFromDate] = React.useState("");
  const [uptoDate, setUptoDate] = React.useState("");
  const [dateError, setDateError] = React.useState("");
  const [isSaving, setIsSaving] = React.useState(false);

  const [selectedUser, setSelectedUser] = React.useState<any>(null);
  const [employeeName, setEmployeeName] = React.useState("");

  const [accessInfo, setAccessInfo] = React.useState<IAccessInfo[]>([]);

  const [filteredData, setFilteredData] = useState<any[]>([]);

  // Pagination
  const itemsPerPage = 50;
  const [currentPage, setCurrentPage] = useState(1);
  const totalPages = Math.ceil(filteredData.length / itemsPerPage);

  // People Picker Context
  const peoplePickerContext: IPeoplePickerContext = {
    absoluteUrl: props.context.pageContext.web.absoluteUrl,
    msGraphClientFactory: props.context.msGraphClientFactory as any,
    spHttpClient: props.context.spHttpClient as any,
  };

  // Load data on page load

  // Fetch List Data
  //   const getAccessInfo = async () => {
  //   const items = await sp.web.lists
  //     .getByTitle("AccrualSheetAccessList")
  //     .items();
  // };

  // React.useEffect(() => {
  //   void getAccessInfo();
  // }, []);
  const getAccessInfo = async () => {
    const items = await sp.web.lists
      .getByTitle("AccrualSheetAccessList")
      .items.select(
        "Title",
        "Division",
        "FromDate",
        "UptoDate",
        "Author/Title",
        "Editor/Title",
        "Created",
        "Modified",
      )
      .expand("Author", "Editor")
      .orderBy("Created", false) // ✅ ADD THIS LINE
      .top(5000)();

    const data = items.map((item: any) => ({
      Division: item.Division,
      EmployeeName: item.Title,
      FromDate: new Date(item.FromDate).toLocaleDateString(),
      UptoDate: new Date(item.UptoDate).toLocaleDateString(),
      CreatedBy: item.Author?.Title,
      CreatedOn: new Date(item.Created).toLocaleDateString(),
      UpdatedBy: item.Editor?.Title,
      UpdatedOn: new Date(item.Modified).toLocaleDateString(),
    }));

    setAccessInfo(data);
    setFilteredData(data);
  };

  const getAccessInfo1 = async () => {
    const items = await sp.web.lists
      .getByTitle("AccrualSheetAccessList")
      .items.select(
        "Title",
        "Division",
        "FromDate",
        "UptoDate",
        "Author/Title",
        "Editor/Title",
        "Created",
        "Modified",
      )
      .expand("Author", "Editor")
      .top(5000)();

    const data = items.map((item: any) => ({
      Division: item.Division,
      EmployeeName: item.Title,
      FromDate: new Date(item.FromDate).toLocaleDateString(),
      UptoDate: new Date(item.UptoDate).toLocaleDateString(),
      CreatedBy: item.Author?.Title,
      CreatedOn: new Date(item.Created).toLocaleDateString(),
      UpdatedBy: item.Editor?.Title,
      UpdatedOn: new Date(item.Modified).toLocaleDateString(),
    }));

    setAccessInfo(data);
  };

  const handleSave1 = async () => {
    if (isSaving) return; // 🚫 prevent double click

    if (!division || !fromDate || !uptoDate) {
      alert("Division, From Date and Upto Date are required");
      return;
    }

    if (uptoDate <= fromDate) {
      alert("Upto Date must be greater than From Date");
      return;
    }

    try {
      setIsSaving(true); // ✅ disable button

      // CASE 1 : User selected
      if (selectedUser) {
        const user = await sp.web.ensureUser(selectedUser.loginName);

        await sp.web.lists.getByTitle("AccrualSheetAccessList").items.add({
          Title: employeeName,
          Division: division,
          FromDate: fromDate,
          UptoDate: uptoDate,
          UsernameId: user.Id,
          Status: "Active",
        });
      }

      // CASE 2 : Group users
      else {
        const users = await sp.web.siteGroups
          .getByName("AccrualUploadAccess")
          .users();

        for (const u of users) {
          await sp.web.lists.getByTitle("AccrualSheetAccessList").items.add({
            Title: u.Title,
            Division: division,
            FromDate: fromDate,
            UptoDate: uptoDate,
            UsernameId: u.Id,
            Status: "Active",
          });
        }
      }

      alert("Access provided successfully");

      setFromDate("");
      setUptoDate("");
      setSelectedUser(null);
      setEmployeeName("");

      setPickerKey((prev) => prev + 1);

      void getAccessInfo();
    } catch (error) {
      console.log(error);
      alert("Error saving data");
    } finally {
      setIsSaving(false); // ✅ enable again (optional)
    }
  };

const handleSave = async () => {
  if (isSaving) return;

  if (!division || !fromDate || !uptoDate) {
    alert("Division, From Date and Upto Date are required");
    return;
  }

  if (new Date(uptoDate) <= new Date(fromDate)) {
    alert("Upto Date must be greater than From Date");
    return;
  }

  if (!selectedUser) {
    alert("Please select a User or Group");
    return;
  }

  try {
    setIsSaving(true);

    console.log("Selected User/Group:", selectedUser);

    const selectedText = selectedUser.text?.trim();
    const selectedKey = selectedUser.key?.trim();

    console.log("Selected Text:", selectedText);
    console.log("Selected Key:", selectedKey);

    // =========================================================
    // STEP 1: Check whether selected item is a SharePoint Group
    // =========================================================

    let sharePointGroup: any = null;

    try {
      if (selectedText) {
        sharePointGroup = await sp.web.siteGroups
          .getByName(selectedText)();

        console.log("SharePoint Group Found:", sharePointGroup);
      }
    } catch (groupError) {
      console.log(
        "Selected item is not a SharePoint group:",
        selectedText
      );
    }

    // =========================================================
    // STEP 2: GROUP
    // =========================================================

    if (sharePointGroup) {
      console.log(
        `Getting users from SharePoint group: ${sharePointGroup.Title}`
      );

      const groupUsers = await sp.web.siteGroups
        .getById(sharePointGroup.Id)
        .users();

      console.log("Group Users:", groupUsers);

      if (!groupUsers || groupUsers.length === 0) {
        alert(
          `No users found in '${sharePointGroup.Title}' group`
        );
        return;
      }

      // Add every member to AccrualSheetAccessList
      for (const user of groupUsers) {
        if (!user.Id) {
          continue;
        }

        console.log("Adding group member:", user);

        await sp.web.lists
          .getByTitle("AccrualSheetAccessList")
          .items.add({
            Title: user.Title,
            Division: division,
            FromDate: fromDate,
            UptoDate: uptoDate,
            UsernameId: user.Id,
            Status: "Active",
          });
      }

      alert(
        `${groupUsers.length} users from '${sharePointGroup.Title}' group added successfully`
      );
    }

    // =========================================================
    // STEP 3: USER
    // =========================================================

    else {
      console.log("Processing individual user");

      if (!selectedUser.loginName && !selectedKey) {
        alert("Unable to determine selected user");
        return;
      }

      const loginName =
        selectedUser.loginName || selectedKey;

      console.log("Ensuring user:", loginName);

      const ensuredUser = await sp.web.ensureUser(loginName);

      console.log("Ensured User:", ensuredUser);

      await sp.web.lists
        .getByTitle("AccrualSheetAccessList")
        .items.add({
          Title: selectedText,
          Division: division,
          FromDate: fromDate,
          UptoDate: uptoDate,
          UsernameId: ensuredUser.Id,
          Status: "Active",
        });

      alert("User access provided successfully");
    }

    // =========================================================
    // STEP 4: RESET FORM
    // =========================================================

    setFromDate("");
    setUptoDate("");
    setSelectedUser(null);
    setEmployeeName("");

    setPickerKey((prev) => prev + 1);

    await getAccessInfo();

  } catch (error) {
    console.error("Save Error:", error);
    alert("Error saving data");
  } finally {
    setIsSaving(false);
  }
};
  const handleExit = () => {
    //window.location.href = `${window.location.origin}/sites/SonaFinance/SitePages/Accuralsheet.aspx`;
          const webUrl = props.context.pageContext.web.absoluteUrl;

    window.location.href = `${webUrl}/SitePages/Accuralsheet.aspx`;
  };

  // Save Button

  const th: any = {
    border: "1px solid #ccc",
    padding: "8px",
    textAlign: "center",
  };

  const td: any = {
    border: "1px solid #ccc",
    padding: "8px",
    textAlign: "center",
  };

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
    setDivision("Railway Business");
    //getAccessInfo();
    void getAccessInfo();
  }, []);


  return (
    <div>
      <div className="header">
        <h1>Manage Access</h1>
      </div>
      <div className="PaddAll">
        <div className="row mb-20">
          <div className="col-md-3">
            <label className="font">Division</label>
            <input value={division} className="form-control readonly" />
          </div>
          <div className="col-md-3">
            <label className="font">User Name</label>

            <PeoplePicker
              key={pickerKey}
              context={peoplePickerContext}
              personSelectionLimit={1}
              showtooltip={true}
              ensureUser={true}
              principalTypes={[
                PrincipalType.User,
                PrincipalType.SharePointGroup,
              ]}
              resolveDelay={500}
              onChange={(items) => {
                console.log(items);

                if (items.length > 0) {
                  setSelectedUser(items[0]);
                  setEmployeeName(items[0].text || "");
                }
              }}
            />
            {/* <PeoplePicker
            key={pickerKey}   // 👈 IMPORTANT
            context={peoplePickerContext}
            personSelectionLimit={1}
            showtooltip={true}
            ensureUser={true}
            principalTypes={[PrincipalType.User]}
            resolveDelay={1000}
            onChange={(items) => {
              if (items.length > 0) {
                setSelectedUser(items[0]);
                setEmployeeName(items[0].text || "");
              }
            }}
          /> */}
            {/* <PeoplePicker context={peoplePickerContext} personSelectionLimit={1} showtooltip={true}
              ensureUser={true} principalTypes={[PrincipalType.User]} resolveDelay={1000}
              onChange={(items) => {
                if (items.length > 0) {
                  setSelectedUser(items[0]);
                  setEmployeeName(items[0].text || "");
                }
              }}
            /> */}
          </div>
          <div className="col-md-3">
            <label className="font">From Date</label>
            <input
              type="date"
              className="form-control"
              value={fromDate}
              onChange={(e) => setFromDate(e.target.value)}
            />
          </div>
          <div className="col-md-3">
            <label className="font">Upto Date</label>
            <input
              type="date"
              className="form-control"
              value={uptoDate}
              min={fromDate}
              onChange={(e) => setUptoDate(e.target.value)}
            />
          </div>
        </div>
        <div className="row mb-20">
          <div
            className="col-md-12"
            style={{ margin: "10px", textAlign: "end" }}
          >
            <button
              onClick={handleSave}
              disabled={isSaving}
              style={{
                opacity: isSaving ? 0.5 : 1,
                cursor: isSaving ? "not-allowed" : "pointer",
              }}
            >
              {isSaving ? "Saving..." : "Save"}
            </button>
          </div>
        </div>

        <hr style={{ border: "1px solid red" }} />

        <div className="heading1">
          <label>Access Info</label>
        </div>
        <div className="main-formcontainer">
          <div className="overflow-x-auto">
            <div className="table-vert-scroll">
              <table className="custom-table min-w-full bg-white rounded-2xl shadow-md">
                <thead
                  style={{ backgroundColor: "#3c3e45" }}
                  className="text-white"
                >
                  <tr>
                    <th className="px-4 py-2">S.No</th>
                    <th className="px-4 py-2">Division</th>
                    <th className="px-4 py-2">Employee Name</th>
                    <th className="px-4 py-2">From Date</th>
                    <th className="px-4 py-2">Upto Date</th>
                    <th className="px-4 py-2">Created By</th>
                    <th className="px-4 py-2">Created On</th>
                    <th className="px-4 py-2">Updated By</th>
                    <th className="px-4 py-2">Updated On</th>
                  </tr>
                </thead>
                <tbody>
                  {paginatedData.map((item, index) => (
                    <tr key={index} className="border-t">
                      <td className="px-4 py-2">{index + 1}</td>
                      <td className="px-4 py-2">{item.Division}</td>
                      <td className="px-4 py-2">{item.EmployeeName}</td>
                      <td className="px-4 py-2">{item.FromDate}</td>
                      <td className="px-4 py-2">{item.UptoDate}</td>
                      <td className="px-4 py-2">{item.CreatedBy}</td>
                      <td className="px-4 py-2">{item.CreatedOn}</td>
                      <td className="px-4 py-2">{item.UpdatedBy}</td>
                      <td className="px-4 py-2">{item.UpdatedOn}</td>
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
        </div>
        <div style={{ textAlign: "center", marginTop: "10px" }}>
          <button onClick={handleExit} className="Reject-btn">
            {" "}
            Exit{" "}
          </button>
        </div>
      </div>
    </div>
  );
}
