// vybe-test: kotlin/data_classes/test_generic_data_class_equality_depends_on_payload
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
            val b = Holder(1)
            val c = Holder(2)
            __check((a == b).toString(), "true")
            __check((a == c).toString(), "false")
        }
