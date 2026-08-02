// vybe-test: csharp/csharp_default_interface_methods_deep/default_method_generic_constraint_on_interface
// origin: languages/csharp/tests/csharp/test_csharp_default_interface_methods_deep.rs

void __Check(string got, string want) {
    if (got != want) {
        Console.WriteLine("FAIL: want [" + want + "] got [" + got + "]");
        throw new Exception("assertion failed");
    }
}

interface ICompare<T> where T:System.IComparable<T>{int Cmp(T a,T b)=>a.CompareTo(b);} class S:ICompare<int>{} __Check((new S().Cmp(3,7)).ToString(), "-1");
