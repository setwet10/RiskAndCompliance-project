using RiskComplianceApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace RiskComplianceApp.Controllers
{
    public class AdminController : Controller
    {
        // GET: Admin
        RiskAppEntities dbObj = new RiskAppEntities();

        public ActionResult AdminLayout()
        {
            
                return View();
            

        }
        public ActionResult AdminTRR()
        {
            TRR obj = new TRR();
            using (var db = new RiskAppEntities())
            {
                var activeRecords = dbObj.TRRs.Where(e => e.IsActive).ToList();
                return View(activeRecords);
            }

        }
        public ActionResult AssignTRR(int id)
        {
            var project = dbObj.TRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("AdminTRR");
        }

        public ActionResult AdminNTRR()
        {
           NTRR obj = new NTRR();
            using (var db = new RiskAppEntities())
            {
                var activeRecords = dbObj.NTRRs.Where(e => e.IsActive).ToList();
                return View(activeRecords);
            }

        }
        public ActionResult AssignNTRR(int id)
        {
            var project = dbObj.NTRRs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("AdminNTRR");
        }
        public ActionResult AdminIBCP()
        {
            IBCP obj = new IBCP();
            using (var db = new RiskAppEntities())
            {
                var filteredRecords = dbObj.IBCPs.Where(x => !x.IsActive).ToList();
                return View(filteredRecords);
            }

        }
        public ActionResult AssignIBCP(int id)
        {
            var project = dbObj.IBCPs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("AdminIBCP");
        }
        public ActionResult AdminKRI()
        {
            KRI obj = new KRI();
            using (var db = new RiskAppEntities())
            {
                var activeRecords = dbObj.KRIs.Where(e => e.IsActive).ToList();
                return View(activeRecords);
            }

        }
        public ActionResult AssignKRI(int id)
        {
            var project = dbObj.KRIs.Find(id);
            if (project != null)
            {
                project.IsActive = !project.IsActive; // Set IsActive to false
                dbObj.SaveChanges();
            }
            return RedirectToAction("AdminKRI");
        }

        public ActionResult ApproveIBCP(int id)
        {
            var record = dbObj.IBCPs.Find(id);
            if (record != null)
            {
                record.Status = "Approved"; // Set Status to Approved
                dbObj.SaveChanges();
            }
            return RedirectToAction("AdminIBCP");
        }

    }
}