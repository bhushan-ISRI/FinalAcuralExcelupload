import * as React from 'react';
//import styles from './DataUpload.module.scss';
 
import './Site.scss'
import ManageAccess from './ManageAccess'
import UploadAccrual from './UploadAccrual'
// import AdjustmentReport from './AdjustmentReport';
import AccuralReport from './AccuralReport'
import AdjustmentReport from './AdjustmentReport'
import { sp } from "../DataUploadWebPart";
import { IDataUploadProps } from './IDataUploadProps';
import PerformerCloserAccess from './PerformerCloserAccess';
import FreezedDashboard from './freezeddashboard';
import ManageAccessIcon from '../assets/access.png';
import UploadSheetIcon from '../assets/upload.png';
import PieChartIcon from '../assets/pie-chart.png';
import LineChartIcon from '../assets/line-chart.png';
import ClosureIcon from '../assets/closure.png';
 
export default function AccrualSheet(props: IDataUploadProps) {
 
  const [page, setPage] = React.useState("home");
  const [hasAccess, setHasAccess] = React.useState<boolean | null>(null);
  const [hasManageAccess, setHasManageAccess] = React.useState(false);
  const [isPerformer, setIsPerformer] = React.useState(false);
  const [canUpload, setCanUpload] = React.useState(false);
 
 
  const openUploadPage = async () => {
 
 
    try {
 
      const currentUser = await sp.web.currentUser();
      const today = new Date();
 
      const items: any = await sp.web.lists
        .getByTitle("AccrualSheetAccessList")
        .items
        .select("FromDate", "UptoDate", "Username/EMail")
        .expand("Username")
        .top(5000)();
 
      const access = items.some((item: any) => {
 
        const from = new Date(item.FromDate);
        const upto = new Date(item.UptoDate);
 
        return (
          item.Username?.EMail === currentUser.Email &&
          today >= from &&
          today <= upto
        );
 
      });
 
      if (access) {
 
        setPage("upload");
 
      } else {
 
        alert("You do not have permission to upload accrual sheet for this period.");
 
      }
 
    } catch (error) {
 
      console.log(error);
 
    }
 
  };
  const checkManageAccess = async () => {
    try {
      const groups = await sp.web.currentUser.groups();
 
      console.log("User Groups:", groups); // 🔍 debug
 
      const isManager = groups.some(
        (group: any) => group.Title === "AccrualManagerAccess"
      );
 
      setHasManageAccess(isManager);
 
    } catch (error) {
      console.log(error);
    }
  };
 
 
 
  const checkUserAccess = async () => {
 
    try {
 
      const currentUser = await sp.web.currentUser();
 
      const today = new Date();
      today.setHours(0, 0, 0, 0); // ✅ remove time
 
      const items: any = await sp.web.lists
        .getByTitle("AccrualSheetAccessList")
        .items
        .select("FromDate", "UptoDate", "Username/EMail")
        .expand("Username")
        .top(5000)();
 
      const access = items.some((item: any) => {
 
        const from = new Date(item.FromDate);
        const upto = new Date(item.UptoDate);
 
        // ✅ remove time from both
        from.setHours(0, 0, 0, 0);
        upto.setHours(0, 0, 0, 0);
 
        return (
          item.Username?.EMail === currentUser.Email &&
          today >= from &&
          today <= upto
        );
 
      });
 
      setHasAccess(true);
      setCanUpload(access);
 
    } catch (error) {
 
      console.log(error);
      setHasAccess(false);
 
    }
 
  };
  const checkUserAccess1 = async () => {
 
    try {
 
      const user = await sp.web.currentUser();
 
      const groups = await sp.web.siteUsers
        .getById(user.Id)
        .groups();
 
      const performer = groups.some(
        (g: any) => g.Title === "AccrualPerformer"
      );
 
      setIsPerformer(performer);
 
    } catch (error) {
 
      console.log("Access check error", error);
 
    }
 
  };
  React.useEffect(() => {
    void checkUserAccess1();
    void checkUserAccess();
    void checkManageAccess(); // ✅ ADD THIS LINE
  }, []);
 
  // Loading state
  if (hasAccess === null) {
    return <h3 style={{ padding: "30px" }}>Loading...</h3>;
  }
 
  // No Access
  // if (!hasAccess) {
  //   return (
  //     <div style={{ padding: "40px" }}>
  //       <h2 style={{ color: "red" }}>Access Denied</h2>
  //       <p>You do not have access to this portal.</p>
  //     </div>
  //   );
  // }
 
  if (page === "manage") {
    return <ManageAccess {...props} />;
  }
 
 if (page === "upload") {
  return <UploadAccrual {...props} />;
}
  if (page === "report") {
    return <AccuralReport {...props}/>;
  }

  if (page === "Freezedreport") {
    return <FreezedDashboard  {...props}/>;
  }
 
  if (page === "report1") {
    return <AdjustmentReport  {...props}/>;
  }
 
  if (page === "Clouser") {
    return <PerformerCloserAccess {...props}/>;
  }
 
  return (
    <div>
      <div className='headSheet'>
        <h2>Accrual Sheet</h2>
      </div>
      <section className="hero">
        <div className="overlay"></div>
        <div className="hero-content">
          <div className="card-container">

            {hasManageAccess && (
              <div className="infoCard" onClick={() => setPage("manage")}>
                <div className="cardContent">
                  <div className='cardalin'>
                    <span className='boximage'>
                      <img src={ManageAccessIcon} alt="" width={25} height={25} />
                    </span>
                    <h4>Manage Access</h4>
                  </div>
                </div>
              </div>
            )}

            {canUpload && (
              <div className="infoCard" onClick={() => setPage("upload")}>
                <div className="cardContent">
                  <div className='cardalin'>
                    <span className='boximage'>
                      <img src={UploadSheetIcon} alt="" width={25} height={25} />
                    </span>
                    <h4>Upload Sheet</h4>
                  </div>
                </div>
              </div>
            )}

           {isPerformer && (
 <div className="infoCard" onClick={() => setPage("report1")}>
              <div className="cardContent">
                <div className='cardalin'>
                  <span className='boximage'>
                    <img src={PieChartIcon} alt="" width={25} height={25} />
                  </span>
                  <h4>Adjustment Report</h4>
                </div>
              </div>
            </div>)}

            {canUpload && ( <div className="infoCard" onClick={() => setPage("report")}>
              <div className="cardContent">
                <div className='cardalin'>
                  <span className='boximage'>
                    <img src={LineChartIcon} alt="" width={25} height={25} />
                  </span>
                  <h4>Accrual Report</h4>
                </div>
              </div>
            </div>)}

            {isPerformer && (
 <div className="infoCard" onClick={() => setPage("Freezedreport")}>
              <div className="cardContent">
                <div className='cardalin'>
                  <span className='boximage'>
                    <img src={LineChartIcon} alt="" width={25} height={25} />
                  </span>
                  <h4>freezed approval report</h4>
                </div>
              </div>
            </div>)}

            {hasManageAccess && (
              <div className="infoCard" onClick={() => setPage("Clouser")}>
                <div className="cardContent">
                  <div className='cardalin'>
                    <span className='boximage'>
                      <img src={ClosureIcon} alt="" width={25} height={25} />
                    </span>
                    <h4>Performer Closure</h4>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </section>
    </div>
  );
}
 
 