using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.IO;
using Apitron.PDF.Kit;
using Apitron.PDF.Kit.FixedLayout;
using Apitron.PDF.Kit.FixedLayout.PageProperties;
using Apitron.PDF.Kit.FixedLayout.Resources;
using Apitron.PDF.Kit.FixedLayout.Resources.Fonts;
using Apitron.PDF.Kit.FlowLayout.Content;
using Apitron.PDF.Kit.Styles;
using Apitron.PDF.Kit.Styles.Appearance;
using Font = Apitron.PDF.Kit.Styles.Text.Font;

// ReSharper disable StringLiteralTypo

namespace InvoiceGenerator.App
{
    [SuppressMessage("ReSharper", "ClassNeverInstantiated.Global")]
    internal class Program
    {
        public static void Main(string[] args)
        {
            using var destinationStream = File.Create(@"..\..\..\output\invoice.pdf");
            GenerateInvoice(destinationStream);
        }

        private static readonly List<ProductEntry> Products = new List<ProductEntry>
        {
            new ProductEntry("Description 1", 0.1, 2, 200, 3),
            new ProductEntry("Description 2", 0.2, 3, 300, 4)
        };

        private static readonly CompanyInfo CompanyInfo = new CompanyInfo("Company Info 1");

        private static readonly CompanyInfo CustomerInfo = new CompanyInfo("Customer Info 2");

        ///<summary>
        /// Generates the invoice based on entered data.
        ///</summary>
        ///<param name="stream">Stream to save the resulting pdf into.</param>
        private static void GenerateInvoice(Stream stream)
        {
            // base path for images
            const string imagesPath = @"..\..\..\images";

            // create document and register styles
            var document = new FlowDocument();

            document.StyleManager.RegisterStyle("grid", new Style
            {
                BorderColor = RgbColors.DarkRed,
                Border = new Border(2)
            });
            
            /* style for products table header, assigned via type + class selectors */
            document.StyleManager.RegisterStyle("gridrow.tableHeader", new Style
            {
                Background = RgbColors.LightSlateGray
            });
            
            document.StyleManager.RegisterStyle("gridrow.tableHeader > *", new Style
            {
                BorderBottom = Border.Solid,
                BorderLeft = Border.Solid,
                BorderColor = RgbColors.LightGray
            });
            
            document.StyleManager.RegisterStyle("gridrow.tableHeader > *.first", new Style
            {
                BorderLeft = Border.None
            });

            /* style matching all cells in rows with class "centerAlignedCells" set
                and all cells in rows with class "centerAlignedCell" set */
            document.StyleManager.RegisterStyle("gridrow.centerAlignedCells > *, gridrow > *.centerAlignedCell",
                new Style
                {
                    Align = Align.Center,
                    Margin = new Thickness(0)
                });

            /* style matching all elements in rows with class "leftAlignedCell" set */
            document.StyleManager.RegisterStyle("gridrow > *.leftAlignedCell", new Style
            {
                Align = Align.Left,
                Padding = new Thickness(5, 0, 0, 0)
            });

            /* default style for any cell in any grid row, assigned via type + child selectors, makes it right aligned */
            document.StyleManager.RegisterStyle("gridrow > *", new Style
            {
                Align = Align.Right,
                Padding = new Thickness(5, 0, 5, 0),
                BorderLeft = Border.Solid,
                BorderBottom = Border.Solid,
                BorderColor = RgbColors.LightGray
            });
            
            document.StyleManager.RegisterStyle("gridrow > *.first", new Style
            {
                BorderLeft = Border.None
            });

            // create resource manager and register image resources
            var resourceManager = new ResourceManager();

            var pathLogo = Path.Combine(imagesPath, "logo.png");
            resourceManager.RegisterResource(
                new Apitron.PDF.Kit.FixedLayout.Resources.XObjects.Image("logo", pathLogo, true)
                {
                    Interpolate = true
                });

            // construct page header which includes store logo and the text "Invoice"
            document.PageHeader.Margin = new Thickness(0, 40, 0, 20);
            document.PageHeader.Padding = new Thickness(10, 0, 10, 0);
            document.PageHeader.Height = 120;
            document.PageHeader.Background = RgbColors.LightGray;
            document.PageHeader.LineHeight = 60;
            document.PageHeader.Add(new Image("logo")
            {
                Height = 50, Width = 50, VerticalAlign = VerticalAlign.Middle
            });
            document.PageHeader.Add(new TextBlock("Invoice")
            {
                Display = Display.InlineBlock,
                Align = Align.Right,
                Font = new Font(StandardFonts.CourierBold, 20),
                Color = RgbColors.Black
            });

            // page content section with padding
            var pageSection = new Section
            {
                Padding = new Thickness(20)
            };

            // add company info section
            pageSection.AddItems(CreateInfoSubsections(new[]
            {
                CompanyInfo.Text, "Bill to:\r\n" + CustomerInfo.Text
            }));

            // add horizontal line for visual separation
            pageSection.Add(new Hr
            {
                Padding = new Thickness(0, 20, 0, 20)
            });

            // add products grid        
            pageSection.Add(CreateProductsGrid());

            // add new line after grid
            pageSection.Add(new Br
            {
                Height = 20
            });

            // add page section into document
            document.Add(pageSection);

            // save document to stream
            document.Write(stream, resourceManager, new PageBoundary(Boundaries.A4));
        }

        ///<summary>
        /// Creates several info sections side by side based on given list of strings.
        ///</summary>
        ///<returns>
        /// List of created sections with textual information.
        ///</returns>
        private static IEnumerable<Section> CreateInfoSubsections(IReadOnlyCollection<string> info)
        {
            var createdSections = new List<Section>();
            var width = 100.0 / info.Count;

            foreach (var t in info)
            {
                var section = new Section
                {
                    Width = Length.FromPercentage(width),
                    Display = Display.InlineBlock
                };

                using (var reader = new StringReader(t))
                {
                    string currentLine;

                    while ((currentLine = reader.ReadLine()) != null)
                    {
                        section.Add(new TextBlock(currentLine));
                        section.Add(new Br());
                    }
                }

                createdSections.Add(section);
            }

            return createdSections;
        }

        ///<summary>
        /// Creates products grid.
        ///</summary>
        private static Grid CreateProductsGrid()
        {
            // create grid content element and its define columns
            var productsGrid = new Grid(20, Length.Auto, 40, 50, 60, 60)
            {
                // add header row
                new GridRow(
                    new TextBlock("#") {Class = "first"},
                    new TextBlock("Product"),
                    new TextBlock("Qty."),
                    new TextBlock("Price"),
                    new TextBlock("Disc.(%)"),
                    new TextBlock("Total"))
                {
                    Class = "tableHeader centerAlignedCells"
                }
            };

            var invoiceTotal = 0.0;

            // enumerate the list of products and create grid rows
            foreach (var product in Products)
            {
                var pos = new TextBlock(product.Pos.ToString()) {Class = "centerAlignedCell first"};

                var description = new TextBlock(product.Description)
                {
                    Class = "leftAlignedCell"
                };

                var qty = new TextBlock(product.Qty.ToString())
                {
                    Class = "centerAlignedCell"
                };

                var price = new TextBlock(
                    product.Price.ToString(CultureInfo.InvariantCulture));

                var discount = new TextBlock(
                    product.Discount.ToString(CultureInfo.InvariantCulture));

                var total = new TextBlock(
                    product.Total.ToString(CultureInfo.InvariantCulture));

                productsGrid.Add(new GridRow(pos, description, qty, price, discount, total));
                invoiceTotal += product.Total;
            }

            // append "total" row
            productsGrid.Add(
                new GridRow(
                    new TextBlock("Total(USD)") {ColSpan = 4},
                    new TextBlock(invoiceTotal.ToString(CultureInfo.InvariantCulture)) {ColSpan = 2}
                ));

            return productsGrid;
        }
    }

    internal class ProductEntry
    {
        public ProductEntry(string description, double discount, int pos, double price, int qty)
        {
            Description = description;
            Discount = discount;
            Pos = pos;
            Price = price;
            Qty = qty;
        }

        public int Pos { get; }
        public int Qty { get; }
        public double Price { get; }
        public double Discount { get; }
        public string Description { get; }
        public double Total => Qty * Price;
    }

    internal class CompanyInfo
    {
        public CompanyInfo(string text)
        {
            Text = text;
        }

        public string Text { get; }
    }
}