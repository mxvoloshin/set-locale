using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;

using OfficeOpenXml;
using OfficeOpenXml.Style;

using SetLocale.Client.Web.Entities;
using SetLocale.Client.Web.Models;
using SetLocale.Client.Web.Services;
using SetLocale.Client.Web.Helpers;

namespace SetLocale.Client.Web.Controllers
{
    public class WordController : BaseController
    {
        private readonly IWordService _wordService;
        public WordController(
            IWordService wordService,
            IUserService userService,
            IFormsAuthenticationService formsAuthenticationService)
            : base(userService, formsAuthenticationService)
        {
            _wordService = wordService;
        }

        [HttpGet, AllowAnonymous]
        public async Task<ActionResult> Detail(string id)
        {
            if (string.IsNullOrEmpty(id)) return RedirectToHome();

            var entity = await _wordService.GetByKey(id);           // Entity null kontrolü Test tarafında yapılmadı.
            if (entity == null) return RedirectToHome();

            var model = WordModel.MapEntityToModel(entity);
            return View(model);
        }

        [HttpGet, AllowAnonymous]
        public async Task<ViewResult> All(int id = 0)
        {
            var pageNumber = id;
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            var items = await _wordService.GetWords(pageNumber);
            var list = items.Items.Select(WordModel.MapEntityToModel).ToList();

            var model = new PageModel<WordModel>
            {
                Items = list,
                HasNextPage = items.HasNextPage,
                HasPreviousPage = items.HasPreviousPage,
                Number = items.Number,
                TotalCount = items.TotalCount,
                TotalPageCount = items.TotalPageCount
            };

            return View(model);
        }

        [HttpGet]
        public async Task<ViewResult> NotTranslated(int id = 1)
        {
            var pageNumber = id;
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            var words = await _wordService.GetNotTranslated(pageNumber);
            var list = words.Items.Select(WordModel.MapEntityToModel).ToList();

            var model = new PageModel<WordModel>
            {
                Items = list,
                HasNextPage = words.HasNextPage,
                HasPreviousPage = words.HasPreviousPage,
                Number = words.Number,
                TotalCount = words.TotalCount,
                TotalPageCount = words.TotalPageCount
            };

            return View(model);
        }

        [HttpGet]
        public ViewResult New()
        {
            var model = new WordModel();
            return View(model);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<ActionResult> New(WordModel model)
        {
            if (!model.IsValidForNew())
            {
                model.Msg = "bir sorun oluştu";
                return View(model);
            }

            model.CreatedBy = User.Identity.GetUserId();

            var key = await _wordService.Create(model);
            if (key == null)
            {
                model.Msg = "bir sorun oluştu, daha önce eklenmiş olabilir";
                return View(model);
            }

            return Redirect("/word/detail/" + key);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<JsonResult> Translate(string key, string language, string translation)
        {
            var model = new ResponseModel { Ok = false };

            if (string.IsNullOrEmpty(key)
                || string.IsNullOrEmpty(language)
                || string.IsNullOrEmpty(translation))
            {
                return Json(model, JsonRequestBehavior.DenyGet);
            }

            model.Ok = await _wordService.Translate(key, language, translation);
            return Json(model, JsonRequestBehavior.DenyGet);
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<JsonResult> Tag(string key, string tag)
        {
            var model = new ResponseModel { Ok = false };

            if (string.IsNullOrEmpty(key)
                || string.IsNullOrEmpty(tag))
            {
                return Json(model, JsonRequestBehavior.DenyGet);
            }

            model.Ok = await _wordService.Tag(key, tag);
            return Json(model, JsonRequestBehavior.DenyGet);
        }

        private async Task<string> ExportWordsToExcel()
        {
            var words = await _wordService.GetAll();

            using (var p = new ExcelPackage())
            {
                p.Workbook.Properties.Title = _htmlHelper.LocalizationString("exported_words");

                p.Workbook.Worksheets.Add(_htmlHelper.LocalizationString("exported_words_sheet_name"));
                var workSheet = p.Workbook.Worksheets[1];

                //display table header
                workSheet.Cells[1, 1].Value = _htmlHelper.LocalizationString("key");
                workSheet.Cells[1, 2].Value = _htmlHelper.LocalizationString("description");
                workSheet.Cells[1, 3].Value = _htmlHelper.LocalizationString("tags");
                workSheet.Cells[1, 4].Value = _htmlHelper.LocalizationString("translation_count");
                workSheet.Cells[1, 5].Value = _htmlHelper.LocalizationString("column_header_translation_tr");
                workSheet.Cells[1, 6].Value = _htmlHelper.LocalizationString("column_header_translation_en");
                workSheet.Cells[1, 7].Value = _htmlHelper.LocalizationString("column_header_translation_az");
                workSheet.Cells[1, 8].Value = _htmlHelper.LocalizationString("column_header_translation_cn");
                workSheet.Cells[1, 9].Value = _htmlHelper.LocalizationString("column_header_translation_fr");
                workSheet.Cells[1, 10].Value = _htmlHelper.LocalizationString("column_header_translation_gr");
                workSheet.Cells[1, 11].Value = _htmlHelper.LocalizationString("column_header_translation_it");
                workSheet.Cells[1, 12].Value = _htmlHelper.LocalizationString("column_header_translation_kz");
                workSheet.Cells[1, 13].Value = _htmlHelper.LocalizationString("column_header_translation_ru");
                workSheet.Cells[1, 14].Value = _htmlHelper.LocalizationString("column_header_translation_sp");
                workSheet.Cells[1, 15].Value = _htmlHelper.LocalizationString("column_header_translation_tk");

                //set styling of header
                workSheet.Cells[1, 1, 1, 15].Style.Font.Bold = true;
                workSheet.Cells[1, 1, 1, 15].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;

                //fill table data
                Func<ICollection<Tag>, string> tagsToString = (tags) =>
                {
                    var result = string.Empty;
                    foreach (var tag in tags)
                    {
                        if (string.IsNullOrEmpty(result))
                        {
                            result = tag.Name;
                        }
                        else
                        {
                            result += string.Format(",{0}", tag.Name);
                        }
                    }
                    return result;
                };

                var tagName = string.Empty;
                for (var i = 0; i < words.Count; i++)
                {
                    var row = i + 2;
                    var word = words[i];

                    workSheet.Cells[row, 1].Value = word.Key;
                    workSheet.Cells[row, 2].Value = word.Description;

                    var tags = tagsToString(word.Tags);
                    if (string.IsNullOrEmpty(tagName))
                    {
                        tagName = tags;
                    }

                    workSheet.Cells[row, 3].Value = tags;
                    workSheet.Cells[row, 4].Value = word.TranslationCount;
                    workSheet.Cells[row, 5].Value = word.Translation_TR;
                    workSheet.Cells[row, 6].Value = word.Translation_EN;
                    workSheet.Cells[row, 7].Value = word.Translation_AZ;
                    workSheet.Cells[row, 8].Value = word.Translation_CN;
                    workSheet.Cells[row, 9].Value = word.Translation_FR;
                    workSheet.Cells[row, 10].Value = word.Translation_GR;
                    workSheet.Cells[row, 11].Value = word.Translation_IT;
                    workSheet.Cells[row, 12].Value = word.Translation_KZ;
                    workSheet.Cells[row, 13].Value = word.Translation_RU;
                    workSheet.Cells[row, 14].Value = word.Translation_SP;
                    workSheet.Cells[row, 15].Value = word.Translation_TK;
                }

                for (var i = 1; i <= 15; i++)
                {
                    workSheet.Column(i).AutoFit();
                }

                var fileName = string.Format("{0}-{1}.xlsx", tagName, DateTime.Now.ToString("s").Replace(':', '-').Replace("T", "-"));
                var filePath = string.Format("/public/files/{0}", fileName);
                var mapPath = Server.MapPath(filePath);

                System.IO.File.WriteAllBytes(mapPath, p.GetAsByteArray());

                return fileName;
            }
        }

        [HttpPost, ValidateAntiForgeryToken]
        public async Task<JsonResult> Export()
        {
            var model = new ResponseModel { Ok = false };
            var fileName = await ExportWordsToExcel();

            model.Ok = true;
            model.Result = fileName;

            return Json(model, JsonRequestBehavior.DenyGet);
        }
    }
}