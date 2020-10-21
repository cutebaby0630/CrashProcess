using System;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Extensions.Configuration;
using SqlServerHelper.Core;
using System.Diagnostics;
using System.Collections.Generic;
using OfficeOpenXml;
using System.ComponentModel;

namespace Crash_AE
{
    class Program
    {
        static void Main(string[] args)
        {
            IConfiguration config = new ConfigurationBuilder().AddJsonFile("appsetting.json", optional: true, reloadOnChange: true).Build();
            string filepath = $"{config[$"TargetFile:FilePath"]}";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            DataTabletoExcel dataTabletoExcel = new DataTabletoExcel();
            DataTable Emergency_dt = dataTabletoExcel.GetDataTableFromExcel(filepath, true);
            List<ColumnData> Emergency_list = dataTabletoExcel.GetColumnDataFromDataTable(Emergency_dt);
            GetStringSql getStringSql = new GetStringSql();
            for (int i = 0; i <= Emergency_list.Count; i++)
            {
                getStringSql.EMDTEmergency(i, Emergency_list[i]);
            }
        }

    }
    public class ColumnData
    {
        public string MedicalNoteId { get; set; }
        public string EmergencyNo { get; set; }
        public string IdentityType { get; set; }
        public string IdentityCode { get; set; }
        public string FavorIdentityType { get; set; }
        public string FavorIdnetityCode { get; set; }
        public string IsReferForm { get; set; }
        public string RegistrationTime { get; set; }
        public string IsWound { get; set; }
        public string IsMajorIllness { get; set; }
        public string LastEmergencyBedTxId { get; set; }
        public string EmgDischargeStateType { get; set; }
        public string EmgDischargeStateCode { get; set; }
        public string DischargeTime { get; set; }
        public string DischargeCloseDate { get; set; }
        public string CancelTime { get; set; }
        public string CancelEmpId { get; set; }
        public string IsObservation { get; set; }
        public string NHICopaymentId { get; set; }
        public string CopaymentType { get; set; }
        public string CopaymentTypeCode { get; set; }
        public string DoctorEvaluateLevel { get; set; }
        public string DoctorVisitTime { get; set; }
        public string NHIBenefitTypeId { get; set; }
        public string NHICaseTypeId { get; set; }
        public string HospitalSubsidyType { get; set; }
        public string HospitalSubsidyCode { get; set; }
        public string GovernmentSubsidyType { get; set; }
        public string GovernmentSubsidyCode { get; set; }
        public string MedicalCategoryType { get; set; }
        public string MedicalCategoryCode { get; set; }
        public string ReferFromHospitalId { get; set; }
        public string DiseasesId { get; set; }
        public string VisitingDoctorId { get; set; }
        public string ClinicDoctorId { get; set; }
        public string EmergencyRetrunReasonType { get; set; }
        public string EmergencyRetrunReasonTypeCode { get; set; }
        public string EmergencyRetrunRemark { get; set; }

        public string PlanDischargeDate { get; set; }
        public string DischargeNoticeDate { get; set; }
        public string ObservationStartTime { get; set; }
        public string ObservationStartEmpId { get; set; }
        public string EyeOpeningScale { get; set; }
        public string VerbalResponseScale { get; set; }
        public string MotorResponseScale { get; set; }
        public string DBP { get; set; }
        public string SBP { get; set; }
        public string BloodOxygen { get; set; }
        public string Weight { get; set; }
        public string ManualTriageLevel { get; set; }
        public string TriageTime { get; set; }
        public string TriageEmpId { get; set; }
        public string COConcentration { get; set; }
        public string LastPainAssessmentId { get; set; }
        public string IsTravel { get; set; }
        public string TravelDate { get; set; }
        public string TravelLocation { get; set; }
        public string IsOccupation { get; set; }
        public string OccupationDescription { get; set; }
        public string IsContact { get; set; }
        public string ContactDescription { get; set; }
        public string TravelEndDate { get; set; }
        public string Temperature { get; set; }
        public string Pulse { get; set; }
        public string Breathe { get; set; }
        public string TriageDivisionCode { get; set; }
        public string DivisionId { get; set; }
        public string IsCluster { get; set; }
        public string ClusterDescription { get; set; }



    }
    #region -- ExceltoData --
    public class DataTabletoExcel
    {
        public DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return tbl;
            }
        }

        public List<ColumnData> GetColumnDataFromDataTable(DataTable dt)
        {
            var convertedList = (from rw in dt.AsEnumerable()
                                 select new ColumnData()
                                 {
                                     MedicalNoteId = Convert.ToString(rw["病歷號"]),
                                     IdentityCode = Convert.ToString(rw["身份類別"]),
                                     FavorIdnetityCode = Convert.ToString(rw["優待身份"]),
                                     IsReferForm = Convert.ToString(rw["是否有轉診單"]),
                                     RegistrationTime = Convert.ToString(rw["掛號時間"]),
                                     IsWound = Convert.ToString(rw["是否外傷"]),
                                     IsMajorIllness = Convert.ToString(rw["是否符合重大傷病"]),
                                     DivisionId = Convert.ToString(rw["檢傷科別"]),
                                     LastEmergencyBedTxId = Convert.ToString(rw["床號"]),
                                     EmgDischargeStateCode = Convert.ToString(rw["離部動向"]),
                                     DischargeTime = Convert.ToString(rw["出院時間"]),
                                     DischargeCloseDate = Convert.ToString(rw["出院結帳日期"]),
                                     CancelEmpId = Convert.ToString(rw["退掛人員"]),
                                     CancelTime = Convert.ToString(rw["退掛時間"]),
                                     IsObservation = Convert.ToString(rw["是否留觀"]),
                                     NHICopaymentId = Convert.ToString(rw["健保部分負擔"]),
                                     CopaymentTypeCode = Convert.ToString(rw["部份負擔"]),
                                     DoctorEvaluateLevel = Convert.ToString(rw["醫師評估等級"]),
                                     DoctorVisitTime = Convert.ToString(rw["醫師看診時間"]),
                                     NHIBenefitTypeId = Convert.ToString(rw["給付類別"]),
                                     NHICaseTypeId = Convert.ToString(rw["健保案件分類"]),
                                     HospitalSubsidyCode = Convert.ToString(rw["醫院補助"]),
                                     GovernmentSubsidyCode = Convert.ToString(rw["政府補助"]),
                                     MedicalCategoryCode = Convert.ToString(rw["就醫類別"]),
                                     ReferFromHospitalId = Convert.ToString(rw["轉入院所"]),
                                     DiseasesId = Convert.ToString(rw["主診斷(ICD10)"]),
                                     VisitingDoctorId = Convert.ToString(rw["主治醫師(員編)"]),
                                     ClinicDoctorId = Convert.ToString(rw["看診醫師(員編)"]),
                                     EmergencyRetrunReasonTypeCode = Convert.ToString(rw["72小時返回原因類別"]),
                                     EmergencyRetrunRemark = Convert.ToString(rw["72小時返回備註"]),
                                     TriageDivisionCode = Convert.ToString(rw["檢傷科別"]),
                                     PlanDischargeDate = Convert.ToString(rw["預計出院日期"]),
                                     DischargeNoticeDate = Convert.ToString(rw["通知出院日期"]),
                                     ObservationStartTime = Convert.ToString(rw["留觀開始時間"]),
                                     ObservationStartEmpId = Convert.ToString(rw["啟動留觀人員"]),
                                     EyeOpeningScale = Convert.ToString(rw["意識(E值)"]),
                                     VerbalResponseScale = Convert.ToString(rw["意識(V值)"]),
                                     MotorResponseScale = Convert.ToString(rw["意識(M值)"]),
                                     Temperature = Convert.ToString(rw["體溫"]),
                                     Pulse = Convert.ToString(rw["脈摶(次/分)"]),
                                     Breathe = Convert.ToString(rw["呼吸(次/分)"]),
                                     DBP = Convert.ToString(rw["血壓(舒張壓)"]),
                                     SBP = Convert.ToString(rw["血壓(收縮壓)"]),
                                     BloodOxygen = Convert.ToString(rw["血氧濃度(%)"]),
                                     Weight = Convert.ToString(rw["體重"]),
                                     ManualTriageLevel = Convert.ToString(rw["檢傷分級"]),
                                     TriageTime = Convert.ToString(rw["檢傷時間(24小時制)"]),
                                     TriageEmpId = Convert.ToString(rw["檢傷人員(員編)"]),
                                     COConcentration = Convert.ToString(rw["一氧化碳濃度(%)"]),
                                     LastPainAssessmentId = Convert.ToString(rw["疼痛評估值"]),
                                     IsTravel = Convert.ToString(rw["是否有旅遊史"]),
                                     TravelDate = Convert.ToString(rw["旅遊日期起"]),
                                     TravelLocation = Convert.ToString(rw["旅遊地點"]),
                                     IsOccupation = Convert.ToString(rw["是否有職業史"]),
                                     OccupationDescription = Convert.ToString(rw["職業史說明"]),
                                     IsContact = Convert.ToString(rw["是否有接觸史"]),
                                     ContactDescription = Convert.ToString(rw["接觸史說明"]),
                                     IsCluster = Convert.ToString(rw["是否有群聚史"]),
                                     ClusterDescription = Convert.ToString(rw["群聚史說明"]),
                                     TravelEndDate = Convert.ToString(rw["旅遊日期迄"]),
                                 }).ToList();

            return convertedList;

        }
    }
    #endregion
    #region -string sql-
    public class GetStringSql
    {
        #region --insert into EMDTEmergency --
        public string EMDTEmergency(int i, ColumnData data)
        {
            return $@"INSERT INTO dbo.EMDTEmergency(
                      EmergencyEncounterId,
                      --EmergencyIdentity - column value is auto-generated
                     EmergencyDate,
                     MedicalNoteId,
                     EmergencyNo,
                     ICCardNo,
                     IdentityType,
                     IdentityCode,
                     ICFormatNo,
                     FavorIdentityType,
                     FavorIdnetityCode,
                     IsReferForm,
                     IsFirstVisit,
                     IsPlanCancel,
                     RegistrationTime,
                     IsReferPatient,
                     FirstTriageId,
                     LastTriageId,
                     IsTriageFinish,
                     IsWound,
                     IsRegistration,
                     IsDoctorVisit,
                     IsMajorIllness,
                     IsFullBed,
                     LastEmergencyBedTxId,
                     EmgDischargeStateType,
                     EmgDischargeStateCode,
                     EmgTransferOutType,
                     EmgTransferOutTypeCode,
                     IsEmergencyRetrun,
                     DischargeTime,
                     DischargeCloseDate,
                     GetHereMethodType,
                     GetHereMethodCode,
                     AttendantType,
                     AttendantTypeCode,
                     BelongingsSecureType,
                     BelongingsSecureCode,
                     CancelTime,
                     CancelEmpId,
                     IsAllergy,
                     NursingDescription,
                     IsObservation,
                     IsTOCC,
                     IsOHCA,
                     IsSpecialCase,
                     IsMCI,
                     IsNewborn,
                     IsPublicInfo,
                     IsFromTriage,
                     NHICopaymentId,
                     TriageDivisionType,
                     TriageDivisionCode,
                     DivisionId,
                     CopaymentType,
                     CopaymentTypeCode,
                     MedicalHistoryType,
                     MedicalHistoryCode,
                     MedicalHistoryDescription,
                     PrepareMovementType,
                     PrepareMovementCode,
                     PrepareMovementDescription,
                     OverTimeType,
                     OverTimeTypeCode,
                     OverTimeDescription,
                     DoctorEvaluateLevel,
                     IsSpecialIdentity,
                     SpecialIdentityDescription,
                     LastExaReportTime,
                     LastLabReportTime,
                     DoctorVisitTime,
                     Remark,
                     NHIBenefitTypeId,
                     NHICaseTypeId,
                     HospitalSubsidyType,
                     HospitalSubsidyCode,
                     GovernmentSubsidyType,
                     GovernmentSubsidyCode,
                     MedicalCategoryType,
                     MedicalCategoryCode,
                     MedicalDateTime,
                     ReferralFromType,
                     ReferralFromCode,
                     ReferFromHospitalId,
                     ReferToHospitalId,
                     DiseasesId,
                     VisitingDoctorId,
                     ClinicDoctorId,
                     ReferralReasonType,
                     ReferralReasonCode,
                     IsNotified,
                     FirstManualTriageLevel,
                     FirstSystemTriageLevel,
                     SpecialTreatmentCode1,
                     SpecialTreatmentCode2,
                     SpecialTreatmentCode3,
                     SpecialTreatmentCode4,
                     NewbornAttachMark,
                     ICCardOrganDonationType,
                     ICCardOrganDonationTypeCode,
                     EmergencyRetrunReasonType,
                     EmergencyRetrunReasonTypeCode,
                     PlanDischargeDate,
                     DischargeNoticeDate,
                     AllergyCheckTime,
                     IsClaimed,
                     CreateTime,
                     CreateEmpId,
                     ModifyTime,
                     ModifyEmpId,
                     OldReceiptNo,
                     IsIsolation,
                     ObservationStartTime,
                     ObservationStartEmpId,
                     EmergencyRetrunRemark
                    )
                    VALUES
                    (
                    (CONVERT([bigint],CONVERT([char](8),dbo.UDF_MingoDateToCEDate(${DateTime.Now}),(112))+'300000000')+(row_number() OVER (ORDER BY INP_CSNO))+@maxIdentity), -- EmergencyEncounterId - bigint
                    -- EmergencyIdentity - int
                    ${DateTime.Now.ToString("yyyy-MM-dd")}, -- EmergencyDate - datetime
                    (SELECT p.MedicalNoteId FROM dbo.PROMMedicalNote p WHERE p.MedicalNoteNo = '${data.MedicalNoteId}'), -- MedicalNoteId - int
                    '', -- EmergencyNo - varchar
                    NULL, -- ICCardNo - varchar
                    1, -- IdentityType - int
                    (select CodeNo FROM promcodes WHERE codetype = '1' AND codename = '${data.IdentityCode}'), -- IdentityCode - varchar
                    null, -- ICFormatNo - varchar
                    16, -- FavorIdentityType - int
                    (select CodeNo FROM promcodes WHERE codetype = '16' AND codename = '${data.FavorIdnetityCode}'), -- FavorIdnetityCode - varchar
                    0, -- IsReferForm - bit
                    0, -- IsFirstVisit - bit true:1 false:0
                    0, -- IsPlanCancel - bit
                    ${DateTime.Now.ToString("yyyy-MM-dd")}, -- RegistrationTime - datetime
                    null, -- IsReferPatient - bit
                    null, -- FirstTriageId - bigint
                    null, -- LastTriageId - bigint
                    null, -- IsTriageFinish - bit
                    '', -- IsWound - bit
                    null, -- IsRegistration - bit
                    null, -- IsDoctorVisit - bit
                    0, -- IsMajorIllness - bit
                    0, -- IsFullBed - bit
                    '', -- LastEmergencyBedTxId - bigint
                    1003, -- EmgDischargeStateType - int
                    (select CodeNo FROM promcodes WHERE codetype = '1003' AND codename = '${data.EmgDischargeStateCode}'), -- EmgDischargeStateCode - varchar
                    1004, -- EmgTransferOutType - int
                    '-9999', -- EmgTransferOutTypeCode - varchar
                    0, -- IsEmergencyRetrun - bit
                    Now(), -- DischargeTime - datetime
                    NULL, -- DischargeCloseDate - date
                    209, -- GetHereMethodType - int
                    '-9999', -- GetHereMethodCode - varchar
                    210, -- AttendantType - int
                    '-9999', -- AttendantTypeCode - varchar
                    211, -- BelongingsSecureType - int
                    '-9999', -- BelongingsSecureCode - varchar
                    null, -- CancelTime - datetime
                    null, -- CancelEmpId - int
                    null, -- IsAllergy - bit
                    null, -- NursingDescription - nvarchar
                    null, -- IsObservation - bit
                    null, -- IsTOCC - bit
                    null, -- IsOHCA - bit
                    null, -- IsSpecialCase - bit
                    0, -- IsMCI - bit
                    0, -- IsNewborn - bit
                    0, -- IsPublicInfo - bit
                    1, -- IsFromTriage - bit
                    (select CodeNo FROM promcodes WHERE codetype = '6032' AND codename = '${data.NHICopaymentId}'), -- NHICopaymentId - int
                    212, -- TriageDivisionType - int
                    (select CodeNo FROM promcodes WHERE codetype = '212' AND codename = '${data.TriageDivisionCode}'), -- TriageDivisionCode - varchar
                    (SELECT * FROM PROMDivision WHERE DivisionCode = ''), -- DivisionId - int
                    32, -- CopaymentType - int
                    (select CodeNo FROM promcodes WHERE codetype = '32' AND codename = '${data.CopaymentTypeCode}'), -- CopaymentTypeCode - varchar
                    213, -- MedicalHistoryType - int
                    '-9999', -- MedicalHistoryCode - varchar
                    NULL, -- MedicalHistoryDescription - nvarchar
                    253, -- PrepareMovementType - int
                    '-9999', -- PrepareMovementCode - varchar
                    null, -- PrepareMovementDescription - nvarchar
                    254, -- OverTimeType - int
                    '-9999', -- OverTimeTypeCode - varchar
                    NULL, -- OverTimeDescription - nvarchar
                    NULL, -- DoctorEvaluateLevel - char
                    0, -- IsSpecialIdentity - bit
	                NULL, -- SpecialIdentityDescription - nvarchar
                    NULL, -- LastExaReportTime - datetime
                    NULL, -- LastLabReportTime - datetime
                    NULL, -- DoctorVisitTime - datetime
                    NULL, -- Remark - nvarchar
                    (select CodeNo FROM promcodes WHERE codetype = '13' AND codename = '${data.NHIBenefitTypeId}'), -- NHIBenefitTypeId - int
                    (select CodeNo FROM promcodes WHERE codetype = '15' AND codename = '${data.NHICaseTypeId}'), -- NHICaseTypeId - int
                    5114, -- HospitalSubsidyType - int
                    (select CodeNo FROM promcodes WHERE codetype = '5114' AND codename = '${data.HospitalSubsidyCode}'), -- HospitalSubsidyCode - varchar
                    5113, -- GovernmentSubsidyType - int
                    (select CodeNo FROM promcodes WHERE codetype = '5113' AND codename = '${data.GovernmentSubsidyCode}'), -- GovernmentSubsidyCode - varchar
                    6011, -- MedicalCategoryType - int
                    (select CodeNo FROM promcodes WHERE codetype = '6011' AND codename = '${data.MedicalCategoryCode}'), -- MedicalCategoryCode - varchar
                    '', -- MedicalDateTime - varchar
                    12, -- ReferralFromType - int
                    '-9999', -- ReferralFromCode - varchar
                    (SELECT ph.ReferHospitalId FROM dbo.PROMReferHospital ph WHERE ph.ReferHospitalCode = '${data.ReferFromHospitalId}'), -- ReferFromHospitalId - int
                    NULL, -- ReferToHospitalId - int
                    (SELECT p.DiseasesId FROM dbo.PROMDiseases p WHERE p.DiseasesCode = '${data.DiseasesId}' ), -- DiseasesId - int
                    (SELECT p.EmpId FROM dbo.PROMEmployee p WHERE p.EmpNo ='${data.VisitingDoctorId}'), -- VisitingDoctorId - int
                    (SELECT p.EmpId FROM dbo.PROMEmployee p WHERE p.EmpNo ='${data.ClinicDoctorId}'), -- ClinicDoctorId - int
                    208, -- ReferralReasonType - int
                    '-9999', -- ReferralReasonCode - varchar
                    0, -- IsNotified - bit
                    0, -- FirstManualTriageLevel - smallint
                    0, -- FirstSystemTriageLevel - smallint
                    '', -- SpecialTreatmentCode1 - char
                    '', -- SpecialTreatmentCode2 - char
                    '', -- SpecialTreatmentCode3 - char
                    '', -- SpecialTreatmentCode4 - char
                    '', -- NewbornAttachMark - char
                    2028, -- ICCardOrganDonationType - int
                    '-9999', -- ICCardOrganDonationTypeCode - varchar
                    2029, -- EmergencyRetrunReasonType - int
                    (select CodeNo FROM promcodes WHERE codetype = '2029' AND codename = '${data.EmergencyRetrunReasonTypeCode}'), -- EmergencyRetrunReasonTypeCode - varchar
                    '2020-10-13 13:55:56', -- PlanDischargeDate - datetime
                    '2020-10-13 13:55:56', -- DischargeNoticeDate - datetime
                    '2020-10-13 13:55:56', -- AllergyCheckTime - datetime
                    0, -- IsClaimed - bit
                    '${DateTime.Now.ToString("yyyy-MM-dd")}', -- CreateTime - datetime
                    0, -- CreateEmpId - int
                    '${DateTime.Now.ToString("yyyy-MM-dd")}', -- ModifyTime - datetime
                    0, -- ModifyEmpId - int
                    '', -- OldReceiptNo - varchar
                    0, -- IsIsolation - bit
                    '${DateTime.Now.ToString("yyyy-MM-dd")}', -- ObservationStartTime - datetime
                    0, -- ObservationStartEmpId - int
                    N'' -- EmergencyRetrunRemark - nvarchar
)";
        }
        #endregion
        #region
        #endregion
    }
    #endregion

}
