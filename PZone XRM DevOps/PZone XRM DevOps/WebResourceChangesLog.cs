using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Xrm.Sdk;
using Microsoft.Xrm.Sdk.Messages;
using Microsoft.Xrm.Sdk.Metadata;
using PZone.Xrm.Plugins;


namespace PZone.Xrm.DevOps
{
    /// <summary>
    /// Запись в журнал (сущность devops_resource_log) информации о действиях над Веб-ресурсами.
    /// </summary>
    /// <remarks>
    /// Плагин делает запись в журнал с информацияе о новых и старых значениях свйоств веб-ресурса,
    /// а так же создает вложения со старыми и новыми значениями веб-ресурса.
    /// </remarks>
    /// <remarks>
    /// <para>Подключение:
    /// <code>
    /// Step
    /// Message:    	        Create
    /// Primary Entity:         webresource
    /// Name:                   Support: Web Resource Create Log
    /// Execution Order:        10
    /// Stage:                  Post-operation
    /// Description:            Журналирование добавления нового веб-ресурса.
    /// 
    /// Step
    /// Message:        	    Update
    /// Primary Entity: 	    webresource
    /// Filtering Attributes:   content, displayname
    /// Name:                   Support: Web Resource Changes Log
    /// Execution Order:        10
    /// Stage:          	    Post-operation
    /// Description:    	    Журналирование обновления веб-ресурса.
    ///         Image
    ///         Image Type:     Pre Image
    ///         Name/Alias:     Image
    ///         Parameters:     webresourcetype, name, displayname, content
    /// 
    /// Step
    /// Message:        	    Delete
    /// Primary Entity: 	    webresource
    /// Name:                   Support: Web Resource Delete Log
    /// Execution Order:        10
    /// Stage:          	    Post-operation
    /// Description:    	    Журналирование обновления веб-ресурса.
    ///         Image
    ///         Image Type:     Pre Image
    ///         Name/Alias:     Image
    ///         Parameters:     webresourcetype, name, displayname, content
    /// </code>
    /// </para>
    /// </remarks>
    // ReSharper disable once RedundantExtendsListEntry
    // ReSharper disable once UnusedMember.Global
    public class WebResourceChangesLog : PluginBase, IPlugin
    {
        private static OptionMetadata[] _webResourceTypes;
        private static Dictionary<Message, OptionSetValue> _actions;
        private static OptionSetValue _webResourceOptionSetValue;


        public WebResourceChangesLog(string unsecureConfig) : base(unsecureConfig)
        {
        }


        public override void Configuring(Context context)
        {
            if (_webResourceTypes == null)
            {
                var response = (RetrieveAttributeResponse)context.Service.Execute(new RetrieveAttributeRequest { EntityLogicalName = "webresource", LogicalName = "webresourcetype" });
                _webResourceTypes = ((PicklistAttributeMetadata)response.AttributeMetadata).OptionSet.Options.ToArray();
            }

            if (_actions == null)
            {
                var response = (RetrieveAttributeResponse)context.Service.Execute(new RetrieveAttributeRequest { EntityLogicalName = "devops_resource_log", LogicalName = "devops_actioncode" });
                var actions = ((PicklistAttributeMetadata)response.AttributeMetadata).OptionSet.Options;
                _actions = new Dictionary<Message, OptionSetValue>();
                var value = actions.FirstOrDefault(a => a.Value % 10 == 1)?.Value;
                if (value.HasValue)
                    _actions.Add(Message.Create, new OptionSetValue(value.Value));
                value = actions.FirstOrDefault(a => a.Value % 10 == 2)?.Value;
                if (value.HasValue)
                    _actions.Add(Message.Update, new OptionSetValue(value.Value));
                value = actions.FirstOrDefault(a => a.Value % 10 == 3)?.Value;
                if (value.HasValue)
                    _actions.Add(Message.Delete, new OptionSetValue(value.Value));
            }

            if (_webResourceOptionSetValue == null)
            {
                var response = (RetrieveAttributeResponse)context.Service.Execute(new RetrieveAttributeRequest { EntityLogicalName = "devops_resource_log", LogicalName = "devops_typecode" });
                _webResourceOptionSetValue = new OptionSetValue(((PicklistAttributeMetadata)response.AttributeMetadata).OptionSet.Options.FirstOrDefault(t => t.Value % 10 == 1)?.Value ?? 0);
            }
        }


        public override IPluginResult Execute(Context context)
        {
            var resource = context.Message != Message.Delete ? context.Entity : new Entity();

            // Если ничего существенно не изменяется, то ничего не делаем.
            if (!resource.Contains("displayname") && !resource.Contains("content") && context.Message != Message.Delete)
                return Ok("Ничего не изменилось.");

            var preResource = context.Message == Message.Create ? new Entity() : context.PreEntityImage;

            var typeAttr = resource.Contains("webresourcetype")
                ? resource.GetAttributeValue<OptionSetValue>("webresourcetype")
                : preResource.GetAttributeValue<OptionSetValue>("webresourcetype");
            var type = typeAttr?.Value ?? 0;

            var fileName = GetFileName(resource.Contains("name")
                ? resource.GetAttributeValue<string>("name")
                : preResource.GetAttributeValue<string>("name"));

            var oldDisplayName = preResource.GetAttributeValue<string>("displayname");
            var displayName = resource.Contains("displayname")
                ? resource.GetAttributeValue<string>("displayname")
                : preResource.GetAttributeValue<string>("displayname");

            var oldEncodedContent = preResource.GetAttributeValue<string>("content");
            var oldContent = Base64Decode(oldEncodedContent);
            var encodedContent = resource.GetAttributeValue<string>("content");
            var content = Base64Decode(encodedContent);

            var mimeType = GetMimeType(type);
            var entity = CreateEntity(context.Service, context, type, fileName, oldDisplayName, displayName, oldContent, content);
            var attachmentFileName = GetAttachmentFileName(type, fileName);
            if (context.Message == Message.Update || context.Message == Message.Delete)
                CreateAttachment(context.Service, entity, "Старый файл", attachmentFileName, mimeType, oldEncodedContent);

            return Ok();
        }


        private static Entity CreateEntity(IOrganizationService service, Context context, int type, string fileName, string oldDisplayName, string displayName, string oldContent, string content)
        {
            var typeName = _webResourceTypes.FirstOrDefault(t => t.Value == type)?.Label?.UserLocalizedLabel?.Label ?? "Unknown";
            var actionNames = new Dictionary<Message, string> { { Message.Create, "Создание" }, { Message.Update, "Обновление" }, { Message.Delete, "Удаление" } };
            var oldVersion = GetVersion(type, oldContent);
            var version = GetVersion(type, content);
            var entity = new Entity("devops_resource_log")
            {
                ["devops_name"] = $"{actionNames[context.Message]} веб-ресурса {typeName} \"{(string.IsNullOrWhiteSpace(displayName) ? fileName : displayName)}\"",
                ["devops_logical_name"] = fileName
            };
            if (_webResourceOptionSetValue.Value != 0)
                entity["devops_typecode"] = _webResourceOptionSetValue; // Web Resource
            if (_actions.ContainsKey(context.Message))
                entity["devops_actioncode"] = _actions[context.Message];
            if (!string.IsNullOrEmpty(oldDisplayName))
                entity["devops_old_display_name"] = oldDisplayName;
            if (!string.IsNullOrEmpty(displayName))
                entity["devops_new_display_name"] = displayName;
            if (!string.IsNullOrEmpty(oldVersion))
                entity["devops_old_version"] = oldVersion;
            if (!string.IsNullOrEmpty(version))
                entity["devops_new_version"] = version;
            entity["devops_object_id"] = context.SourceContext.PrimaryEntityId.ToString();
            entity.Id = service.Create(entity);
            return entity;
        }


        private static void CreateAttachment(IOrganizationService service, Entity entity, string subject, string fileName, string mimeType, string encodedContent)
        {
            var annotationOld = new Entity("annotation")
            {
                ["objectid"] = entity.ToEntityReference(),
                ["subject"] = subject,
                ["documentbody"] = encodedContent,
                ["mimetype"] = mimeType,
                ["filename"] = fileName
            };
            service.Create(annotationOld);
        }


        private static string GetFileName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return "unknown";
            var index = name.LastIndexOf("/", StringComparison.Ordinal);
            return index > 0 ? name.Substring(index + 1) : name;
        }


        /// <summary>
        /// Изменение расширения для запрещенных по умолчанию типов файлов.
        /// </summary> 
        /// <param name="type">Тип файла.</param>
        /// <param name="fileName">Имя файла.</param>
        /// <returns>
        /// Метод возвращает новое имя файла, которое можно использовать в примечаних.
        /// </returns>
        /// <remarks>
        /// Файлы некоторых типов по умолчанию не разрешено сохранять в качестве вложений. Поэтому 
        /// для таких файлов расширение изменяется на *.txt.
        /// </remarks>
        private static string GetAttachmentFileName(int type, string fileName)
        {
            var templates = new Dictionary<int, string>
            {
                //{ 1, @"text/html" }, // Веб-страница (HTML)
                //{ 2, @"text/css" }, // Таблица стилей (CSS)
                { 3, @"{0}.txt" }, // Скрипт (JScript)
                //{ 4, @"text/xml" }, // Данные (XML)
                //{ 5, @"image/png" }, // Формат PNG
                //{ 6, @"image/jpeg" }, // Формат JPG
                //{ 7, @"image/gif" }, // Формат GIF
                //{ 8, @"application/octet-stream" }, // Silverlight (XAP)
                //{ 9, @"application/xml" }, // Таблица стилей (XSL)
                //{ 10, @"image/x-icon" } // Формат ICO
            };
            return templates.ContainsKey(type) ? string.Format(templates[type], fileName) : fileName;
        }


        /// <exclude />
        public static string GetVersion(int resourceTypeValue, string resourceBody)
        {
            var patterns = new Dictionary<int, string>
            {
                { 1, "<html[^>]+data-version=\"(?<version>[^\"]+)\"" }, // Веб-страница (HTML)
                { 2, "\\/\\*[\\*!][\\s\\S]*\\* @version (?<version>[\\d+\\.]*)\\s*$[\\s\\S]*\\*\\/" }, // Таблица стилей (CSS)
                { 3, "\\/\\*[\\*!][\\s\\S]*\\* @version (?<version>[\\d+\\.]*)\\s*$[\\s\\S]*\\*\\/" } // Скрипт (JScript)
            };
            if (!patterns.ContainsKey(resourceTypeValue) || string.IsNullOrWhiteSpace(resourceBody))
                return string.Empty;

            var pattern = patterns[resourceTypeValue];
            var regex = new Regex(pattern, RegexOptions.Multiline);
            var versionNumber = regex.Match(resourceBody);
            return versionNumber.Success ? versionNumber.Groups["version"].Value : string.Empty;
        }


        private static string GetMimeType(int resourceTypeValue)
        {
            var mimeTypes = new Dictionary<int, string>
            {
                { 1, @"text/html" }, // Веб-страница (HTML)
                { 2, @"text/css" }, // Таблица стилей (CSS)
                { 3, @"application/javascript" }, // Скрипт (JScript)
                { 4, @"text/xml" }, // Данные (XML)
                { 5, @"image/png" }, // Формат PNG
                { 6, @"image/jpeg" }, // Формат JPG
                { 7, @"image/gif" }, // Формат GIF
                { 8, @"application/octet-stream" }, // Silverlight (XAP)
                { 9, @"application/xml" }, // Таблица стилей (XSL)
                { 10, @"image/x-icon" } // Формат ICO
            };
            return mimeTypes.ContainsKey(resourceTypeValue) ? mimeTypes[resourceTypeValue] : string.Empty;
        }


        public static string Base64Decode(string base64EncodedData)
        {
            if (string.IsNullOrWhiteSpace(base64EncodedData))
                return string.Empty;
            var base64EncodedBytes = Convert.FromBase64String(base64EncodedData);
            return Encoding.UTF8.GetString(base64EncodedBytes);
        }
    }
}