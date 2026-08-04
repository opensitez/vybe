// vybe-test: java/varargs/varargs_prints_each_argument_in_order
// origin: languages/java/tests/java/test_varargs.rs

public class Main {
static void show(int... nums) {
            for (int n : nums) System.out.println(n);
        }
    public static void main(String[] args) {
show(4, 5, 6);
    }
}

