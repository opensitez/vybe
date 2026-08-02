// vybe-test: csharp/csharp_with_expression_records/with_class_record_equal_same_values
// origin: languages/csharp/tests/csharp/test_csharp_with_expression_records.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

record Node(int Id); var a=new Node(1); var b=a with{Id=1}; __Check((a==b).ToString(), "True");
