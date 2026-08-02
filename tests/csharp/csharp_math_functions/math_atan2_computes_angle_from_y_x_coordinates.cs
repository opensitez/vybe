// vybe-test: csharp/csharp_math_functions/math_atan2_computes_angle_from_y_x_coordinates
// origin: languages/csharp/tests/csharp/test_csharp_math_functions.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

double angle = System.Math.Atan2(1, 1);
__Check((System.Math.Round(angle, 4)).ToString(), "0.7854");
