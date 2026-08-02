// vybe-test: csharp/csharp_string_span/span_copy_to_writes_into_destination
// origin: languages/csharp/tests/csharp/test_csharp_string_span.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

int[] src={1,2,3};
int[] dst=new int[3];
src.AsSpan().CopyTo(dst);
__Check((dst[2]).ToString(), "3");
