// vybe-test: csharp/csharp_pattern_relational/switch_expression_relational_greater_arm_grade_b
// origin: languages/csharp/tests/csharp/test_csharp_pattern_relational.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int score=85; __Check((score switch{>=90=>"A",>=80=>"B",_=>"C"}).ToString(), "B");
