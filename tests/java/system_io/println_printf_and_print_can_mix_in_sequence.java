// vybe-test: java/system_io/println_printf_and_print_can_mix_in_sequence
// origin: languages/java/tests/java/test_system_io.rs

public class Main {
    public static void main(String[] args) {
System.out.print("["); System.out.printf("%d", 5); System.out.println("]");
    }
}

