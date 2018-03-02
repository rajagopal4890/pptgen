using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace DocumentService.Models
{
    public class Inspection
    {
        public int Id { get; set; }

        public double TimeTakenInSec { get; set; }

        public string ManagerAssignedForReview { get; set; }

        public string LotNumber { get; set; }

        public int? LotQuantity { get; set; }

        public int? SampleQuantity { get; set; }

        public string RevisionNumberAndDate { get; set; }

        public string PartNo { get; set; }

        public string SupplierName { get; set; }

        public DateTime DateCreated { get; set; }

        public string UserCreated { get; set; }

        public InspectionStatus Status { get; set; }
    }

    public enum InspectionStatus
    {
        Inspecting = 1,
        Waiting_For_Approval = 2,
        Part_Accepted = 3,
        Part_Rejected = 4,
        Recheck = 5
    }
}