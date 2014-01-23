using System;
using System.Collections.Generic;
using System.Data.Entity.Migrations;
using System.Linq;

using SetLocale.Client.Web.Entities;
using SetLocale.Client.Web.Helpers;

namespace SetLocale.Client.Web.Repositories
{
    public class SetLocaleDbMigrationConfiguration : DbMigrationsConfiguration<SetLocaleDbContext>
    {
        public SetLocaleDbMigrationConfiguration()
        {
            AutomaticMigrationsEnabled = true;
            AutomaticMigrationDataLossAllowed = true;
        }

        protected override void Seed(SetLocaleDbContext context)
        {
            if (context.Users.Any()) return;

            #region Users

            AddAdmin(context, "Serdar Büyüktemiz", "hserdarb@gmail.com");
            AddAdmin(context, "Caner Çavuş", "canercvs@gmail.com");
            AddAdmin(context, "Ramiz Sümer", "ramiz.sumerr@gmail.com");
            AddAdmin(context, "Mehmet Sabancıoğlu", "mehmet.sabancioglu@gmail.com");
            AddAdmin(context, "Cihan Çoşkun", "cihancoskun@gmail.com");
            AddAdmin(context, "Kemal Çolak", "kml.colak@gmail.com");
            AddAdmin(context, "Duygu Sevim", "sevimduygu@gmail.com");

            AddTranslator(context, "Translator", "translator@test.com");

            #endregion

            #region Apps

            AddApplication(context, "setlocale@test.com", "set-locale", "a localization management application.",
                "setlocale.com");
            AddApplication(context, "setcrm@test.com", "set-crm", "a brand new crm application.", "setcrm.com");
            AddApplication(context, "setmembership@test.com", "set-membership", "a membership management application.",
                "setmembership.com");
            AddApplication(context, "drone@test.com", "Marmara Drone",
                "a wireless control dashboard for unmanned air vehicle.", "github.com/jupre/marmaradrone");
            AddApplication(context, "collade@test.com", "Collade",
                "a task management and team collaboration application.", "collade.com");

            #endregion

            #region Words
            AddWord(context, "menu_words", string.Empty, "Kelimeler", "Words", "Слова", "set-locale");
            AddWord(context, "menu_words_words", string.Empty, "Tüm Kelimeler", "All Words", "Все Слова", "set-locale");
            AddWord(context, "menu_words_my_words", string.Empty, "Kelimelerim", "My Words", "Мои Слова", "set-locale");
            AddWord(context, "menu_words_new_word", string.Empty, "Yeni Kelime", "New Word", "Новое Слово", "set-locale");
            AddWord(context, "menu_words_not_translated", string.Empty, "Çevrilmeyen Kelime", "Word Not Translated", "Слово Без Перевода", "set-locale");
            AddWord(context, "menu_apps", string.Empty, "Uygulamalar", "Applications", "Приложения", "set-locale");
            AddWord(context, "menu_apps_apps", string.Empty, "Uygulamalarım", "My Applications", "Мои Приложения", "set-locale");
            AddWord(context, "menu_apps_new_app", string.Empty, "Yeni Uygulama", "New Application", "Новое Приложение", "set-locale");
            AddWord(context, "menu_settings", string.Empty, "Yönetim", "Administration", "Администратор", "set-locale");
            AddWord(context, "menu_settings_apps", string.Empty, "Tüm Uygulamalar", "All Applications", "Все Приложения", "set-locale");
            AddWord(context, "menu_settings_users", string.Empty, "Tüm Kullanıcılar", "All Users", "Все Пользователи", "set-locale");
            AddWord(context, "menu_settings_new_translator", string.Empty, "Yeni Çevirmen", "New Translator", "Новый Переводчик", "set-locale");
            AddWord(context, "menu_user_login", string.Empty, "Giriş", "Login", "Вход", "set-locale");
            AddWord(context, "menu_user_logout", string.Empty, "Çıkış", "Logout", "Выход", "set-locale");
            AddWord(context, "menu_user_sign_up", string.Empty, "Kayıt Ol", "Sign Up", "Войти", "set-locale");
            AddWord(context, "menu_user_reset", string.Empty, "Şifre Sıfırla", "Reset Password", "Сбросить Пароль", "set-locale");
            AddWord(context, "menu_search", string.Empty, "Ara", "Search", "Поиск", "set-locale");

            AddWord(context, "email", string.Empty, "E-posta", "Email", "Почта", "set-locale");
            AddWord(context, "name", string.Empty, "İsim", "Name", "Имя", "set-locale");
            AddWord(context, "app_owner_email", string.Empty, "Uygulama Sahibinin E-Postası", "Application of the E-mail owners", "Приложение Пользователей Почты", "set-locale");
            AddWord(context, "password", string.Empty, "Şifre", "Password", "Пароль", "set-locale");
            AddWord(context, "app_name", string.Empty, "Uygulama İsmi", "Application Name", "Название Приложения", "set-locale");
            AddWord(context, "description", string.Empty, "Açıklama", "Description", "Описание", "set-locale");
            AddWord(context, "usage_count", string.Empty, "Kullanım Sayısı", "Usage Count", "Количество Использования", "set-locale");
            AddWord(context, "url", string.Empty, "Url", "Url", "Url", "set-locale");
            AddWord(context, "token", string.Empty, "Token", "Token", "Признак", "set-locale");
            AddWord(context, "cretion_date", string.Empty, "Oluşturma Tarihi", "Creation Date", "Дата Создания", "set-locale");
            AddWord(context, "user_role", string.Empty, "Yetki Grubu", "Role", "Роль", "set-locale");
            AddWord(context, "word_key", string.Empty, "Anahtar", "Key", "Ключь", "set-locale");
            AddWord(context, "tag", string.Empty, "Etiket", "Tag", "Тег", "set-locale");
            AddWord(context, "translated_language", string.Empty, "Çevrilmiş Dil", "Translated Language", "Язык Перевода", "set-locale");
            AddWord(context, "translation", string.Empty, "Çeviri", "Translation", "Перевод", "set-locale");
            AddWord(context, "language", string.Empty, "Dil", "Language", "Язык", "set-locale");
            AddWord(context, "forgot_your_password", string.Empty, "Şifremi Unuttum", "Forgot My Password", "Забыл Пароль", "set-locale");
            AddWord(context, "total_page_count", string.Empty, "Toplam Sayfa Sayısı", "Total Page Count", "Всего Страниц", "set-locale");
            AddWord(context, "translator_name", string.Empty, "İsim", "Name", "Название", "set-locale");

            AddWord(context, "btn_login", string.Empty, "Giriş", "Login", "Вход", "set-locale");
            AddWord(context, "btn_sign_up", string.Empty, "Kayıt Ol", "Sign Up", "Выход", "set-locale");
            AddWord(context, "btn_password_reset", string.Empty, "Şifre Sıfırlama Linki Gönder", "Send Password Reset Link", "Отпарвить Ссылку Для Сброса Пароля", "set-locale");
            AddWord(context, "btn_create_new_token", string.Empty, "Yeni Token Oluştur", "Create New Token", "Новый Признак", "set-locale");
            AddWord(context, "btn_new_word", string.Empty, "Yeni Çeviri Ekle", "Add New Translation", "Добавить Перевод", "set-locale");
            AddWord(context, "btn_deactivate", string.Empty, "Pasif", "Deactivate", "Деактивировать", "set-locale");
            AddWord(context, "btn_activate", string.Empty, "Aktif", "Activate", "Активировать", "set-locale");
            AddWord(context, "btn_delete", string.Empty, "Sil", "Delete", "Удалить", "set-locale");
            AddWord(context, "btn_save", string.Empty, "Kaydet", "Save", "Сохранить", "set-locale");
            AddWord(context, "btn_edit", string.Empty, "Düzenle", "Edit", "Редактировать", "set-locale");
            AddWord(context, "btn_cancel", string.Empty, "İptal", "Cancel", "Отменить", "set-locale");
            AddWord(context, "btn_ok", string.Empty, "Tamam", "Ok", "Да", "set-locale");
            AddWord(context, "btn_export_to_excel", string.Empty, "Excel", "Excel", "Excel", "set-locale");
            AddWord(context, "column_header_translation_tr", string.Empty, "Tükçe", "Turkish", "Турецкий","set-locale");
            AddWord(context, "column_header_translation_az", string.Empty, "Azerbeycan", "Azerbaijan", "Азербайджанский","set-locale");
            AddWord(context, "column_header_translation_cn", string.Empty, "Çince", "Chinese", "Китайский","set-locale");
            AddWord(context, "column_header_translation_fr", string.Empty, "Fransızca", "Français", "Французский","set-locale");
            AddWord(context, "column_header_translation_gr", string.Empty, "Yunanca", "Greek", "Греческий","set-locale");
            AddWord(context, "column_header_translation_it", string.Empty, "İtalyanca", "İtaliano", "Итальянский","set-locale");
            AddWord(context, "column_header_translation_kz", string.Empty, "Kazakça", "Kazakh", "Казахский","set-locale");
            AddWord(context, "column_header_translation_ru", string.Empty, "Rusça", "Russian", "Русский","set-locale");
            AddWord(context, "column_header_translation_sp", string.Empty, "İspanyolca", "Espanol", "Испанский","set-locale");
            AddWord(context, "column_header_translation_tk", string.Empty, "Türkmençe", "Turkic", "Турецкий","set-locale");


            AddWord(context, "home_title", string.Empty, "Set-locale Servisine Hoş Geldiniz.", "Welcome To Set-locale.", "Добро Пожаловать В Set-locale.", "set-locale");
            AddWord(context, "words_key_listing_title", string.Empty, "Anahtar Listesi", "Key Listing", "Список Ключей","set-locale");
            AddWord(context, "login_view_title", string.Empty, "Sisteme Giriş", "İnto The System", "Вход В Систему","set-locale");
            AddWord(context, "new_user_title", string.Empty, "Yeni Kullanıcı", "New User", "Новый пользователь","set-locale");
            AddWord(context, "password_reset_title", string.Empty, "Şifre Sıfırla", "Reset Password", "Сбросить Пароль","set-locale");
            AddWord(context, "user_apps_title", string.Empty, "Uygulamalarım", "My Applications", "Мои приложения","set-locale");
            AddWord(context, "all_apps_title", string.Empty, "Tüm Uygulamalar", "All Applications", "Все приложения","set-locale");
            AddWord(context, "all_users_title", string.Empty, "Tüm Kullanıcılar", "All Users", "Все пользователи","set-locale");
            AddWord(context, "word_new_key_title", string.Empty, "Yeni Anahtar", "New Key", "Новый Ключь","set-locale");
            AddWord(context, "menu_words_my_words_key_listing_title", string.Empty, "Anahtar Listesi", "My Key Listing", "Мой Список Ключей","set-locale");
            AddWord(context, "new_app_title", string.Empty, "Yeni Uygulama", "New Application", "Новое Приложение","set-locale");
            AddWord(context, "new_translator_title", string.Empty, "Yeni Çevirmen", "New Translator", "Новый Переводчик","set-locale");
            AddWord(context, "tag_keys_title", string.Empty, "Tag Anahtar Listesi", "Tag Keys Listing", "Тег Списка Ключей","set-locale");

            AddWord(context, "modal_title_apps", string.Empty, "Uygulama Durumu", "Application Status", "Статус Приложения","set-locale");
            AddWord(context, "modal_body_apps", string.Empty, "Değiştirmek İstediğinize Emin misiniz?", "Are You Sure Want To Change The Application's Status?", "Вы Уверены Что Хотите Изменить Статус Приложения?","set-locale");
            AddWord(context, "modal_title_delete_token", string.Empty, "Token Sil", "Delete Token", "Удалить Признак","set-locale");
            AddWord(context, "modal_body_delete_token", string.Empty, "Bu Anahtarı Silmek İstediğinize Emin misiniz?", "Are You Sure You Want To Delete This Token?", "Вы Уверены Что Хотите Удалить Этот Признак?","set-locale");
            AddWord(context, "modal_title_users", string.Empty, "Kullanıcı Durumu", "User Status", "Статус Пользователя","set-locale");
            AddWord(context, "modal_body_users", string.Empty, "Durumu Değiştirmek İstediğinize Emin misiniz?", "Are You Sure You Want To Change The User's Status?", "Вы Уверены Что Хотите Изменить Статус Пользователя?","set-locale"); 
            #endregion
        }

        private static void AddAdmin(SetLocaleDbContext context, string name, string email)
        {
            var user = new User
            {
                Email = email,
                Name = name,
                RoleId = SetLocaleRole.Admin.Value,
                RoleName = SetLocaleRole.Admin.ToString(),
                ImageUrl = GravatarHelper.GetGravatarURL(email, 35),
                PasswordHash = BCrypt.Net.BCrypt.HashPassword("password"),
                LastLoginAt = DateTime.Now,
                IsActive = true
            };
            context.Users.Add(user);
        }

        private static void AddTranslator(SetLocaleDbContext context, string name, string email)
        {
            var user = new User
            {
                Email = email,
                Name = name,
                RoleId = SetLocaleRole.Translator.Value,
                RoleName = SetLocaleRole.Translator.ToString(),
                ImageUrl = GravatarHelper.GetGravatarURL(email, 35),
                PasswordHash = BCrypt.Net.BCrypt.HashPassword("password"),
                LastLoginAt = DateTime.Now,
                IsActive = true
            };
            context.Users.Add(user);
        }

        private static void AddApplication(SetLocaleDbContext context, string email, string name, string description, string url)
        {
            var app = new App
            {
                UserEmail = email,
                Name = name,
                Description = description,
                Url = url,
                CreatedBy = 1,
                Tokens = new List<Token>
                {
                    new Token
                    {
                        Key = Guid.NewGuid().ToString().Replace("-", string.Empty),
                        UsageCount = 0,
                        CreatedBy = 1,
                        IsAppActive = true
                    }
                },
                IsActive = true
            };

            context.Apps.Add(app);
        }

        private static void AddWord(SetLocaleDbContext context, string key, string description, string tr, string en, string ru, string tagName)
        {
            var word = new Word
            {
                Key = key,
                Description = description,
                IsTranslated = true,
                Translation_EN = en,
                Translation_TR = tr,
                Translation_RU = ru,
                CreatedBy = 1,
                TranslationCount = 2,
                Tags = new List<Tag>
                {
                    new Tag
                    {
                        Name = tagName,
                        UrlName = tagName.ToUrlSlug(),
                        CreatedBy = 1
                    }
                }
            };
            context.Words.Add(word);
        }
    }
}