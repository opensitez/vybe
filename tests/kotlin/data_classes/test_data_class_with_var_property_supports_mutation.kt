// vybe-test: kotlin/data_classes/test_data_class_with_var_property_supports_mutation
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Counter(var value: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val c = Counter(3)
            c.value += 4
            __check((c.value).toString(), "7")
        }
