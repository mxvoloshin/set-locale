using System;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using OfficeOpenXml;
using SetLocale.Client.Web.Entities;
using SetLocale.Client.Web.Helpers;
using SetLocale.Client.Web.Models;
using SetLocale.Client.Web.Services;

namespace SetLocale.Client.Web.Controllers
{
    public class AdminController : BaseController
    {
        private readonly IAppService _appService;
        private readonly IWordService _wordService;

        public AdminController(
            IUserService userService,
            IWordService wordService,
            IFormsAuthenticationService formsAuthenticationService, 
            IAppService appService)
            : base(userService, formsAuthenticationService)
        {
            _appService = appService;
            _wordService = wordService;
        }

        protected override void OnActionExecuting(ActionExecutingContext filterContext)
        {
            //todo: commented because no admin user exists

            /*if (CurrentUser.RoleId != SetLocaleRole.Admin.Value)
            {
                filterContext.Result = RedirectToHome();
            }*/

            
            base.OnActionExecuting(filterContext); 
        }

        [HttpGet]
        public ActionResult Index()
        {
            return Redirect("/admin/apps");
        }

        [HttpGet]
        public ViewResult NewTranslator()
        {
            var model = new UserModel();
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<ActionResult> NewTranslator(UserModel model)
        {
            if (!model.IsValidForNewTranslator())
            {
                model.Msg = "bir sorun oluştu...";
                return View(model);
            }

            model.Password = Guid.NewGuid().ToString().Replace("-", string.Empty);
            model.Language = Thread.CurrentThread.CurrentUICulture.Name;
            var userId = await _userService.Create(model, SetLocaleRole.Translator.Value);
            if (userId == null)
            {
                model.Msg = "bir sorun oluştu...";
                return View(model);
            }

            //send mail to translator to welcome and ask for reset password

            return Redirect("/admin/users");
        }
         
        [HttpGet]
        public async Task<ActionResult> Users(int id = 0, int page = 1)
        {
            var pageNumber = page;
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            PagedList<User> users;

            ViewBag.RoleId = id;
            if (SetLocaleRole.IsValid(id))
            {
                users = await _userService.GetAllByRoleId(id, pageNumber);
            }
            else
            {
                users = await _userService.GetUsers(pageNumber);
            }

            var list = users.Items.Select(UserModel.MapUserToUserModel).ToList();

            var model = new PageModel<UserModel>
            {
                Items = list,
                HasNextPage = users.HasNextPage,
                HasPreviousPage = users.HasPreviousPage,
                Number = users.Number,
                TotalCount = users.TotalCount,
                TotalPageCount = users.TotalPageCount
            };
             
            return View(model);
        }
          
        [HttpGet]
        public async Task<ActionResult> Apps(int id = 0)
        {
            var pageNumber = id; 
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }
            var apps = await _appService.GetApps(pageNumber);

            var list = apps.Items.Select(AppModel.MapFromEntity).ToList();

            var model = new PageModel<AppModel>
            {
                Items = list,
                HasNextPage = apps.HasNextPage,
                HasPreviousPage = apps.HasPreviousPage,
                Number = apps.Number,
                TotalCount = apps.TotalCount,
                TotalPageCount = apps.TotalPageCount
            };

            return View(model);
        }

        [HttpGet, AllowAnonymous]
        public ActionResult Import()
        {
            var model = new AppModel();
            return View(model);
        }

        [HttpPost, AllowAnonymous]
        public async Task<ActionResult> Import(AppModel model, HttpPostedFileBase file)
        {
            //check file format
            var fileInfo = new FileInfo(file.FileName);
            
            if (!String.Equals(fileInfo.Extension, ".xlsx"))
            {
                model.Msg = "File format is incorrect, *.xlsx file expected";
                return View(model);
            }

            //read file data in buffer
            var buf = new byte[file.InputStream.Length];
            file.InputStream.Read(buf, 0, (int)file.InputStream.Length);
            
            var memoryStream = new MemoryStream(buf);

            using (var p = new ExcelPackage(memoryStream))
            {
                if (!p.Workbook.Worksheets.Any())
                {
                    model.Msg = "No Worksheets in selected file";
                    return View(model);
                }
                
                var workSheet = p.Workbook.Worksheets[1];

                //check header of the excel file for correct format
                string errMessage;
                if (!CheckExcelFileColumns(workSheet, out errMessage))
                {
                    model.Msg = errMessage;
                    return View(model);
                }

                //read file data by row
                var addedWordsCount = 0;
                var skipedWordsCount = 0;
                var rowNum = 2;

                while (rowNum <= workSheet.Dimension.End.Row)
                {

                    var word = await CreateWordFromRow(workSheet, rowNum);
                    rowNum++;

                    var wordModel = WordModel.MapEntityToModel(word);
                    wordModel.CreatedBy = User.Identity.GetUserId();

                    var addedWord = _wordService.Create(wordModel);
                    if (addedWord == null)
                        skipedWordsCount++;
                    else
                        addedWordsCount++;
                }

                model.Msg = String.Format("Added new words: {0}.", addedWordsCount);
                if (skipedWordsCount > 0)
                    model.Msg = string.Format("{0} Words skipped: {1}.", model.Msg, skipedWordsCount);
            }

            return View(model);
        }

        private bool CheckExcelFileColumns(ExcelWorksheet sheet, out string errMessage)
        {
            errMessage = string.Empty;

            try
            {
                Action<ExcelWorksheet, int, int, string> checkCellValue = (worksheet, row, column, checkString) =>
                {
                    var cellValue = worksheet.Cells[row, column].Value;
                    var result = string.Equals(cellValue, _htmlHelper.LocalizationString(checkString));
                    if (!result)
                        throw new ArgumentException("Column unrecognized.", checkString);
                };

                checkCellValue(sheet, 1, 1, "key");
                checkCellValue(sheet, 1, 2, "description");
                checkCellValue(sheet, 1, 3, "tags");
                checkCellValue(sheet, 1, 4, "translation_count");
                checkCellValue(sheet, 1, 5, "column_header_translation_tr");
                checkCellValue(sheet, 1, 6, "column_header_translation_en");
                checkCellValue(sheet, 1, 7, "column_header_translation_az");
                checkCellValue(sheet, 1, 8, "column_header_translation_cn");
                checkCellValue(sheet, 1, 9, "column_header_translation_fr");
                checkCellValue(sheet, 1, 10, "column_header_translation_gr");
                checkCellValue(sheet, 1, 11, "column_header_translation_it");
                checkCellValue(sheet, 1, 12, "column_header_translation_kz");
                checkCellValue(sheet, 1, 13, "column_header_translation_ru");
                checkCellValue(sheet, 1, 14, "column_header_translation_sp");
                checkCellValue(sheet, 1, 15, "column_header_translation_tk");

                return true;
            }
            catch (ArgumentException arg)
            {
                errMessage = arg.Message;
                return false;
            }
        }

        private async Task<Word> CreateWordFromRow(ExcelWorksheet sheet, int rowNumber)
        {
            var newWord = new Word();

            newWord.Key = sheet.Cells[rowNumber, 1].Value.ToString();
            newWord.Description = sheet.Cells[rowNumber, 2].Value == null ? string.Empty : sheet.Cells[rowNumber, 2].Value.ToString();
            
            //parse tags
            var tags = sheet.Cells[rowNumber, 3].Value == null ? string.Empty : sheet.Cells[rowNumber, 3].Value.ToString();
            var tagsArray = tags.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            
            newWord.Tags = new Collection<Tag>();
            foreach (var item in tagsArray)
            {
                newWord.Tags.Add(new Tag
                {
                    Name = item,
                    UrlName = item.ToUrlSlug()
                });   
            }

            //set translations
            newWord.TranslationCount = Convert.ToInt32(sheet.Cells[rowNumber, 4].Value);
            newWord.Translation_TR = sheet.Cells[rowNumber, 5].Value == null ? string.Empty : sheet.Cells[rowNumber, 5].Value.ToString();
            newWord.Translation_EN = sheet.Cells[rowNumber, 6].Value == null ? string.Empty : sheet.Cells[rowNumber, 6].Value.ToString();
            newWord.Translation_AZ = sheet.Cells[rowNumber, 7].Value == null ? string.Empty : sheet.Cells[rowNumber, 7].Value.ToString();
            newWord.Translation_CN = sheet.Cells[rowNumber, 8].Value == null ? string.Empty : sheet.Cells[rowNumber, 8].Value.ToString();
            newWord.Translation_FR = sheet.Cells[rowNumber, 9].Value == null ? string.Empty : sheet.Cells[rowNumber, 9].Value.ToString();
            newWord.Translation_GR = sheet.Cells[rowNumber, 10].Value == null ? string.Empty : sheet.Cells[rowNumber, 10].Value.ToString();
            newWord.Translation_IT = sheet.Cells[rowNumber, 11].Value == null ? string.Empty : sheet.Cells[rowNumber, 11].Value.ToString();
            newWord.Translation_KZ = sheet.Cells[rowNumber, 12].Value == null ? string.Empty : sheet.Cells[rowNumber, 12].Value.ToString();
            newWord.Translation_RU = sheet.Cells[rowNumber, 13].Value == null ? string.Empty : sheet.Cells[rowNumber, 13].Value.ToString();
            newWord.Translation_SP = sheet.Cells[rowNumber, 14].Value == null ? string.Empty : sheet.Cells[rowNumber, 14].Value.ToString();
            newWord.Translation_TK = sheet.Cells[rowNumber, 15].Value == null ? string.Empty : sheet.Cells[rowNumber, 15].Value.ToString();

            newWord.IsTranslated = true;

            return newWord;
        }
    }
}