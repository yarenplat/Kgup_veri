using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;
using OfficeOpenXml;

class Program
{
    static void Main()
    {
        // JSON verisini bir dosyadan oku veya string olarak elde et
        string jsonVeri = File.ReadAllText("C:\\Users\\YarenPOLAT\\Desktop\\response.json");

        // JSON verisini C# nesnesine dönüştür
        var veriNesnesi = JsonConvert.DeserializeObject<VeriNesnesi>(jsonVeri);

        // Excel dosyası oluştur
        using (var package = new ExcelPackage())
        {
            // Excel sayfası oluştur
            var worksheet = package.Workbook.Worksheets.Add("VeriSayfasi");

            // Başlık satırını ekleyerek başla
            YazdirBasliklar(worksheet, 1, 1);

            // JSON verisini Excel'e yaz
            YazdirVeri(veriNesnesi.items, worksheet, 2, 1);

            // Excel dosyasını kaydet
            package.SaveAs(new FileInfo("sonuc.xlsx"));
        }

        Console.WriteLine("İşlem tamamlandı. Excel dosyası oluşturuldu.");
    }

    static void YazdirBasliklar(ExcelWorksheet worksheet, int satir, int sutun)
    {
        // Excel sayfasına başlık satırını yaz
        var basliklar = new List<string>
        {
            "Date", "Time", "Toplam", "Dogalgaz", "Ruzgar", "Linyit",
            "TasKomur", "IthalKomur", "FuelOil", "Jeotermal", "Barajli", "Nafta",
            "Biokutle", "Akarsu", "Diger"
        };

        for (int i = 0; i < basliklar.Count; i++)
        {
            worksheet.Cells[satir, sutun + i].Value = basliklar[i];
        }
    }

    static void YazdirVeri(List<Item> veri, ExcelWorksheet worksheet, int satir, int sutun)
    {
        // Her bir JSON öğesini Excel satırına yaz
        foreach (var item in veri)
        {
            worksheet.Cells[satir, sutun].Value = item.Date;
            worksheet.Cells[satir, sutun + 1].Value = item.Time;
            worksheet.Cells[satir, sutun + 2].Value = item.Toplam;
            worksheet.Cells[satir, sutun + 3].Value = item.Dogalgaz;
            worksheet.Cells[satir, sutun + 4].Value = item.Ruzgar;
            worksheet.Cells[satir, sutun + 5].Value = item.Linyit;
            worksheet.Cells[satir, sutun + 6].Value = item.TasKomur;
            worksheet.Cells[satir, sutun + 7].Value = item.IthalKomur;
            worksheet.Cells[satir, sutun + 8].Value = item.FuelOil;
            worksheet.Cells[satir, sutun + 9].Value = item.Jeotermal;
            worksheet.Cells[satir, sutun + 10].Value = item.Barajli;
            worksheet.Cells[satir, sutun + 11].Value = item.Nafta;
            worksheet.Cells[satir, sutun + 12].Value = item.Biokutle;
            worksheet.Cells[satir, sutun + 13].Value = item.Akarsu;
            worksheet.Cells[satir, sutun + 14].Value = item.Diger;


            satir++;
        }
    }
}

public class Item
{
    public DateTime Date { get; set; }
    public string Time { get; set; }
    public double Toplam { get; set; }
    public double Dogalgaz { get; set; }
    public double Ruzgar { get; set; }
    public double Linyit { get; set; }
    public double TasKomur { get; set; }
    public double IthalKomur { get; set; }
    public double FuelOil { get; set; }
    public double Jeotermal { get; set; }
    public double Barajli { get; set; }
    public double Nafta { get; set; }
    public double Biokutle { get; set; }
    public double Akarsu { get; set; }
    public double Diger { get; set; }
}

public class Page
{
    public int number { get; set; }
    public int size { get; set; }
    public int total { get; set; }
    public Sort sort { get; set; }
}

public class Sort
{
    public string field { get; set; }
    public string direction { get; set; }
}

public class Totals
{
    public DateTime Date { get; set; }
    public string Time { get; set; }
    public double Toplam { get; set; }
    public double Dogalgaz { get; set; }
    public double Ruzgar { get; set; }
    public double Linyit { get; set; }
    public double TasKomur { get; set; }
    public double IthalKomur { get; set; }
    public double FuelOil { get; set; }
    public double Jeotermal { get; set; }
    public double Barajli { get; set; }
    public double Nafta { get; set; }
    public double Biokutle { get; set; }
    public double Akarsu { get; set; }
    public double Diger { get; set; }
}

public class VeriNesnesi
{
    public List<Item> items { get; set; }
    public Page page { get; set; }
    public Totals totals { get; set; }
}
