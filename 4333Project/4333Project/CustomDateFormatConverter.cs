using Newtonsoft.Json;
using System;
using System.Globalization;

public class CustomDateFormatConverter : JsonConverter
{
    private const string DateFormat = "dd.MM.yyyy"; // Формат даты

    public override bool CanConvert(Type objectType)
    {
        return objectType == typeof(DateTime); // Конвертируем только в DateTime
    }

    public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
    {
        var dateString = reader.Value as string;

        if (dateString != null)
        {
            // Попытка распарсить дату в формате "dd.MM.yyyy"
            DateTime date;
            if (DateTime.TryParseExact(dateString, DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            {
                return date; // Если парсинг успешен, возвращаем дату
            }
        }

        // В случае ошибки распарсить дату
        throw new JsonException($"Invalid date format: {reader.Value}");
    }

    public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
    {
        // Здесь не будем менять формат даты при сериализации (это не требуется)
        writer.WriteValue(((DateTime)value).ToString(DateFormat));
    }
}
