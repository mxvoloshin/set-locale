using System;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using Moq;
using NUnit.Framework;

using SetLocale.Client.Web.Entities;
using SetLocale.Client.Web.Helpers;
using SetLocale.Client.Web.Repositories;
using SetLocale.Client.Web.Test.Builders;

namespace SetLocale.Client.Web.Test.Services
{
    [TestFixture]
    public class TagServiceTests
    {
        [Test]
        public async void should_set_pagenumber_to_one_if_it_is_less()
        {
            //arrange
            var wordRepository = new Mock<IRepository<Word>>();
            wordRepository.Setup(x => x.FindAll(It.IsAny<Expression<Func<Word, bool>>>()))
                          .Returns(new List<Word> { new Word { Id = 1, Key = "key" } }.AsQueryable());

            //act
            var sut = new TagServiceBuilder().WithWordRepository(wordRepository.Object)
                                             .Build();

            var result = await sut.GetWords(string.Empty, 0);

            //assert
            Assert.AreEqual(result.Number, 1);
        }

        [Test]
        public async void should_return_pagedlist()
        {
            //arrange
            var words = new List<Word>();
            for (var i = 1; i <= ConstHelper.PageSize * 4; i++)
                words.Add(new Word { Id = i });

            var wordRepository = new Mock<IRepository<Word>>();
            wordRepository.Setup(x => x.FindAll(It.IsAny<Expression<Func<Word, bool>>>(), It.IsAny<Expression<Func<Word, object>>>()))
                          .Returns(words.AsQueryable());

            //act
            var sut = new TagServiceBuilder().WithWordRepository(wordRepository.Object)
                                             .Build();

            var result = await sut.GetWords("url", 2);

            //assert
            Assert.AreEqual(result.Number, 2);
            Assert.AreEqual(result.Size, ConstHelper.PageSize);
            Assert.AreEqual(result.TotalCount, ConstHelper.PageSize * 4);
            Assert.AreEqual(result.HasPreviousPage, true);
            Assert.AreEqual(result.HasNextPage, true);
            Assert.AreEqual(result.Items.Count, ConstHelper.PageSize);
            Assert.AreEqual(result.Items.First().Id, (ConstHelper.PageSize * 4) - (ConstHelper.PageSize));
            Assert.AreEqual(result.Items.Last().Id, (ConstHelper.PageSize * 4) - (ConstHelper.PageSize * 2) + 1);
        }
    }
}