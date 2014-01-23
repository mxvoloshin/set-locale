using System;
using System.Threading.Tasks;
using System.Linq;
using System.Collections.Generic;

using SetLocale.Client.Web.Entities;
using SetLocale.Client.Web.Helpers;
using SetLocale.Client.Web.Models;
using SetLocale.Client.Web.Repositories;

namespace SetLocale.Client.Web.Services
{
    public interface IWordService
    {
        Task<string> Create(WordModel model);
        Task<PagedList<Word>> GetByUserId(int userId, int pageNumber);
        Task<Word> GetByKey(string key); 
        Task<PagedList<Word>> GetWords(int pageNumber);
        Task<PagedList<Word>> GetNotTranslated(int pageNumber);
        Task<bool> Translate(string key, string language, string translation);
        Task<bool> Tag(string key, string tag);
        Task<List<Word>> GetAll();
    }

    public class WordService : IWordService
    {
        private readonly IRepository<Word> _wordRepository; 
        public WordService(IRepository<Word> wordRepository)
        {
            _wordRepository = wordRepository;
        }

        public Task<string> Create(WordModel model)
        {
            if (!model.IsValidForNew())
            {
                return null;
            }

            var slug = model.Key.ToUrlSlug();
            if (_wordRepository.Any(x => x.Key == slug))
            {
                return null;
            }

            var tags = new List<Tag>();
            var items = model.Tag.Split(new[] { "," }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var item in items)
            {
                tags.Add(new Tag
                {
                    CreatedBy = model.CreatedBy,
                    Name = item,
                    UrlName = item.ToUrlSlug()
                });
            }

            var word = new Word
            {
                Key = slug,
                Description = model.Description ?? string.Empty,
                IsTranslated = model.IsTranslated,
                TranslationCount = model.Translations.Count,
                CreatedBy = model.CreatedBy,
                UpdatedBy = model.CreatedBy,
                Tags = tags
            };

            if (model.IsTranslated)
            {
                foreach (var item in model.Translations)
                {
                    if (item.Language.Key == LanguageModel.TR().Key)
                        word.Translation_TR = item.Value;
                    else if (item.Language.Key == LanguageModel.EN().Key)
                        word.Translation_EN = item.Value;
                    else if (item.Language.Key == LanguageModel.AZ().Key)
                        word.Translation_AZ = item.Value;
                    else if (item.Language.Key == LanguageModel.CN().Key)
                        word.Translation_CN = item.Value;
                    else if (item.Language.Key == LanguageModel.FR().Key)
                        word.Translation_FR = item.Value;
                    else if (item.Language.Key == LanguageModel.GR().Key)
                        word.Translation_GR = item.Value;
                    else if (item.Language.Key == LanguageModel.IT().Key)
                        word.Translation_IT = item.Value;
                    else if (item.Language.Key == LanguageModel.KZ().Key)
                        word.Translation_KZ = item.Value;
                    else if (item.Language.Key == LanguageModel.RU().Key)
                        word.Translation_RU = item.Value;
                    else if (item.Language.Key == LanguageModel.SP().Key)
                        word.Translation_SP = item.Value;
                    else if (item.Language.Key == LanguageModel.TK().Key)
                        word.Translation_TK = item.Value;
                }
            }

            _wordRepository.Create(word);
            
            if (!_wordRepository.SaveChanges())
                return null;
            
            return Task.FromResult(word.Key);
        }

        public Task<PagedList<Word>> GetByUserId(int userId,int pageNumber)
        {
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            if (userId < 1)
            {
                return null;
            }

            var words = _wordRepository.FindAll(x => x.CreatedBy == userId, x => x.Tags).ToList();

            long totalCount = words.Count();
            var totalPageCount = (int)Math.Ceiling(totalCount / (double)ConstHelper.PageSize);

            if (pageNumber > totalPageCount)
            {
                pageNumber = 1;
            }

            words = words.OrderByDescending(x => x.Id).Skip(ConstHelper.PageSize * (pageNumber - 1)).Take(ConstHelper.PageSize).ToList();

            return Task.FromResult(new PagedList<Word>(pageNumber, ConstHelper.PageSize, totalCount, words));
        }

        public Task<Word> GetByKey(string key)
        {
            var word = _wordRepository.FindOne(x => x.Key == key);
            return Task.FromResult(word);
        }

        public Task<PagedList<Word>> GetWords(int pageNumber)
        {
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            pageNumber--;
  
            var items = _wordRepository.FindAll();

            long totalCount = items.Count();
            var totalPageCount = (int)Math.Ceiling(totalCount / (double)ConstHelper.PageSize);

            if (pageNumber > totalPageCount)
            {
                pageNumber = 1;
            }

            items = items.OrderByDescending(x => x.Id).Skip(ConstHelper.PageSize * (pageNumber)).Take(ConstHelper.PageSize);

            return Task.FromResult(new PagedList<Word>(pageNumber, ConstHelper.PageSize, totalCount, items.ToList()));
        }

        public Task<PagedList<Word>> GetNotTranslated(int pageNumber = 1)
        {
            if (pageNumber < 1)
            {
                pageNumber = 1;
            }

            var words = _wordRepository.FindAll(x => x.IsTranslated == false).ToList();

            long totalCount = words.Count();
            var totalPageCount = (int)Math.Ceiling(totalCount / (double)ConstHelper.PageSize);

            if (pageNumber > totalPageCount)
            {
                pageNumber = 1;
            }

            words = words.OrderByDescending(x => x.Id).Skip(ConstHelper.PageSize * (pageNumber - 1)).Take(ConstHelper.PageSize).ToList();

            return Task.FromResult(new PagedList<Word>(pageNumber, ConstHelper.PageSize, totalCount, words));
        }

        public Task<bool> Translate(string key, string language, string translation)
        {
            if (string.IsNullOrEmpty(key)
               || string.IsNullOrEmpty(language)
               || string.IsNullOrEmpty(translation))
            {
                return Task.FromResult(false);
            }

            var word = _wordRepository.FindOne(x => x.Key == key);
            if (word == null)
            {
                return Task.FromResult(false);
            }

            var type = word.GetType();
            var propInfo = type.GetProperty(string.Format("Translation_{0}", language.ToUpperInvariant()), new Type[0]);
            if (propInfo == null)
            {
                return Task.FromResult(false);
            }

            propInfo.SetValue(word, translation);
            word.TranslationCount++;
            word.IsTranslated = true;

            _wordRepository.Update(word);

            return Task.FromResult(_wordRepository.SaveChanges());
        }

        public Task<bool> Tag(string key, string tagName)
        {
            if (string.IsNullOrEmpty(key)
               || string.IsNullOrEmpty(tagName))
            {
                return Task.FromResult(false);
            }

            var word = _wordRepository.FindOne(x => x.Key == key && x.Tags.Count(y => y.Name == tagName) == 0);
            if (word == null)
            {
                return Task.FromResult(false);
            }

            var tag = new Tag
           {
               Name = tagName,
               UrlName = tagName.ToUrlSlug(),
               CreatedBy = 1
           };

            word.Tags = new List<Tag> { tag };

            _wordRepository.Update(word);

            return Task.FromResult(_wordRepository.SaveChanges());
        }

        public Task<List<Word>> GetAll()
        {
            var words = _wordRepository.FindAll().ToList();

            return Task.FromResult(words);
        }
    }
}