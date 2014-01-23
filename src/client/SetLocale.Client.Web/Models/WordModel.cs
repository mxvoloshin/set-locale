using System.Linq;
using System.Text;
using SetLocale.Client.Web.Entities;
using System.Collections.Generic;

namespace SetLocale.Client.Web.Models
{
    public class WordModel : BaseModel
    {
        public string Key { get; set; }
        public int Count { get; set; }
        public string Description { get; set; }
        public string Tag { get; set; }
        public bool IsTranslated { get; set; }
        public List<TagModel> Tags { get; set; }
        public List<LanguageModel> Languages { get; set; }
        public List<TranslationModel> Translations { get; set; }
        public int CreatedBy { get; set; }
        public string LanguagesStr {
            get
            {
                var str = new StringBuilder();
                foreach (var language in Languages)
                {
                    str.AppendFormat("{{ id:'{1}',text:'{0}' }},", language.Name, language.Key);
                }

                return str.ToString();
            }
        }

        public WordModel()
        {
            Tags = new List<TagModel>();
            Languages = new List<LanguageModel>();
            Translations = new List<TranslationModel>();
        }

        public bool IsValidForNew()
        {
            return !string.IsNullOrEmpty(Key)
                   && !string.IsNullOrEmpty(Tag);
        }

        public static WordModel MapEntityToModel(Word entity)
        {
            var model = new WordModel
            {
                Key = entity.Key,
                Description = entity.Description,
                IsTranslated = entity.IsTranslated,
                CreatedBy = entity.CreatedBy
            };

            if (entity.Tags != null
                && entity.Tags.Any())
            {
                foreach (var tag in entity.Tags)
                {
                    model.Tags.Add(new TagModel
                    {
                        Name = tag.Name,
                        UrlName = tag.UrlName
                    });

                    model.Tag += string.Format("{0},", tag.Name);
                }
            }

            if (entity.IsTranslated)
            {
                if (!string.IsNullOrEmpty(entity.Translation_TR))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_TR,
                        Language = LanguageModel.TR()
                    });

                    model.Languages.Add(LanguageModel.TR());
                }

                if (!string.IsNullOrEmpty(entity.Translation_EN))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_EN,
                        Language = LanguageModel.EN()
                    });

                    model.Languages.Add(LanguageModel.EN());
                }

                if (!string.IsNullOrEmpty(entity.Translation_AZ))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_AZ,
                        Language = LanguageModel.AZ()
                    });

                    model.Languages.Add(LanguageModel.AZ());
                }

                if (!string.IsNullOrEmpty(entity.Translation_SP))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_SP,
                        Language = LanguageModel.SP()
                    });

                    model.Languages.Add(LanguageModel.SP());
                }
                if (!string.IsNullOrEmpty(entity.Translation_CN))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_CN,
                        Language = LanguageModel.CN()
                    });

                    model.Languages.Add(LanguageModel.CN());
                }
                if (!string.IsNullOrEmpty(entity.Translation_FR))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_FR,
                        Language = LanguageModel.FR()
                    });

                    model.Languages.Add(LanguageModel.FR());
                }
                if (!string.IsNullOrEmpty(entity.Translation_GR))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_GR,
                        Language = LanguageModel.GR()
                    });

                    model.Languages.Add(LanguageModel.GR());
                }
                if (!string.IsNullOrEmpty(entity.Translation_IT))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_IT,
                        Language = LanguageModel.IT()
                    });

                    model.Languages.Add(LanguageModel.IT());
                }
                if (!string.IsNullOrEmpty(entity.Translation_KZ))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_KZ,
                        Language = LanguageModel.KZ()
                    });

                    model.Languages.Add(LanguageModel.KZ());
                }
                if (!string.IsNullOrEmpty(entity.Translation_RU))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_RU,
                        Language = LanguageModel.RU()
                    });

                    model.Languages.Add(LanguageModel.RU());
                }
                if (!string.IsNullOrEmpty(entity.Translation_TK))
                {
                    model.Translations.Add(new TranslationModel
                    {
                        Key = entity.Key,
                        Value = entity.Translation_TK,
                        Language = LanguageModel.TK()
                    });

                    model.Languages.Add(LanguageModel.TK());
                }
            }

            return model;
        }
    }
}