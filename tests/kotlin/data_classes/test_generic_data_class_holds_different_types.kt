// vybe-test: kotlin/data_classes/test_generic_data_class_holds_different_types
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Holder<T>(val value: T)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Holder(1)
            val b = Holder("x")
            __check((a.value).toString(), "1")
            __check((b.value).toString(), "x")
        }
