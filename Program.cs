using Microsoft.Exchange.WebServices.Data;

var service = new ExchangeService(ExchangeVersion.Exchange2016, TimeZoneInfo.Local)
{ Url = new Uri("https://exchange.server.ru/EWS/Exchange.asmx") };
service.Credentials = new WebCredentials("username", "Password", "domain");

var inbox = await Folder.Bind(service, WellKnownFolderName.Inbox);
Console.WriteLine("Unread count: {0}", inbox?.UnreadCount);
var inboxView = new ItemView(inbox.UnreadCount)
{
    PropertySet = new PropertySet(BasePropertySet.FirstClassProperties)
};
var itemsResults = await inbox.FindItems(inboxView);
await service.LoadPropertiesForItems(itemsResults.Items, new PropertySet(ItemSchema.Body, ItemSchema.Subject));

foreach (Item item in itemsResults.Items)
{
    var message = await EmailMessage.Bind(service, item.Id, new PropertySet(ItemSchema.Attachments));
    if(message == null)
        { continue; }
    // 2. Перебираем коллекцию вложений
    foreach (Attachment attachment in message.Attachments)
    {
        if (attachment is FileAttachment fileAttachment)
        {
            // Обработка файлового вложения (если нужно)
            await fileAttachment.Load("C:\\temp\\" + fileAttachment.Name);
            Console.WriteLine("Файловое вложение: " + fileAttachment.Name);
        }
        else if (attachment is ItemAttachment itemAttachment)
        {

            // 4. Загружаем вложение в память
            // Этот вызов приводит к запросу GetAttachment к EWS.
            var response = await itemAttachment.Load();

            Console.WriteLine("Вложение-сообщение: " + itemAttachment.Name);

            // 5. Получаем доступ к содержимому прикрепленного сообщения
            // Проверяем, является ли элемент именно письмом
            if (itemAttachment.Item is EmailMessage attachedMessage)
            {
                Console.WriteLine("Тема прикрепленного письма: " + attachedMessage.Subject);
                Console.WriteLine("От кого: " + attachedMessage.From.Address);
                Console.WriteLine("Текст письма: " + attachedMessage.Body.Text);
                // ... работа со свойствами attachedMessage
            }
        }
    }
}
