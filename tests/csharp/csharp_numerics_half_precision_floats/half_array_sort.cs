// vybe-test: csharp/csharp_numerics_half_precision_floats/half_array_sort

string __buf = "";

void __P(string s) {
    __buf = __buf + s + "\n";
}

void __Pr(string s) {
    __buf = __buf + s;
}

void __Check(string want) {
    if (__buf != want && __buf != want + "\n") {
        Console.WriteLine("FAIL: want [" + want + "] got [" + __buf + "]");
        throw new Exception("assertion failed");
    }
}

Half[] arr = new Half[] { (Half)5.0f, (Half)1.0f, (Half)3.0f };
Array.Sort(arr);
__P(arr[0].ToString(System.Globalization.CultureInfo.InvariantCulture));
__P(arr[2].ToString(System.Globalization.CultureInfo.InvariantCulture));
__Check("1\n5");
