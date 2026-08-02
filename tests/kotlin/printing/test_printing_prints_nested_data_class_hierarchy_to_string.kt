// vybe-test: kotlin/printing/test_printing_prints_nested_data_class_hierarchy_to_string
// origin: languages/kotlin/tests/kotlin/test_printing.rs

data class Child(val id: Int)
        data class Wrapper(val child: Child)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Wrapper(Child(7))).toString(), "Wrapper(child=Child(id=7))")
        }
