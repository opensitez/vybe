// vybe-test: csharp/winforms/control_create_graphics_returns_graphics_handle
// origin: languages/csharp/tests/csharp/test_winforms.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

var pb = new PictureBox();
        pb.Name = "art";
        var g = pb.CreateGraphics();
        __Check((g == null ? "null-graphics" : "have-graphics").ToString(), "have-graphics");
