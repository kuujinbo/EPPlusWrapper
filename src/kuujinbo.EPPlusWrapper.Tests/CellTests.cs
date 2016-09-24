using Xunit;
using System.Drawing;

namespace kuujinbo.EPPlusWrapper.Tests
{
    public class CellTests
    {
        [Fact]
        public void GetHtmlColor_NamedColor_ReturnsColor()
        {
            Assert.Equal(Color.Blue, Cell.GetHtmlColor("blue"));
        }

        [Fact]
        public void GetHtmlColor_HexColor_ReturnsColor()
        {
            var color = Cell.GetHtmlColor("#000000");
            var black = Color.Black;

            Assert.Equal(black.A, color.A);
            Assert.Equal(black.R, color.R);
            Assert.Equal(black.G, color.G);
            Assert.Equal(black.B, color.B);
        }
    }
}