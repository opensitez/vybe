// vybe-test: java/print_stream_format/print_stream_append_then_printf_on_same_line
// origin: languages/java/tests/java/test_print_stream_format.rs

public class Main {
    public static void main(String[] args) {
System.out.append("["); System.out.printf("%s", "x"); System.out.println("]");
    }
}

