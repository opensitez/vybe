// vybe-test: java/overloading/overload_void_print_int_vs_string
// origin: languages/java/tests/java/test_overloading.rs

public class Main {
static void print(int n) { System.out.println("i" + n); }
        static void print(String s) { System.out.println("s" + s); }
    public static void main(String[] args) {
print(3); print("ok");
    }
}

