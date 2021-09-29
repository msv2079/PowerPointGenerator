using System;
using System.IO;
using System.Linq;
using PowerPointGenerator;

namespace TestPowerPointGenerator
{
    static class Program
    {
        static void Main()
        {
            var powerPointGenerator = new PowerPointFactory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Teste.pptx"));
            var imageFilePaths = Directory.GetFiles(@"..\..\..\Imagem", "*.jpg").ToList();
            powerPointGenerator.CreateTitleAndImageSlides(imageFilePaths);
        }
    }
}
