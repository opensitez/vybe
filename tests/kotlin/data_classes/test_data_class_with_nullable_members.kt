// vybe-test: kotlin/data_classes/test_data_class_with_nullable_members
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Holder(val value: String?)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val missing = Holder(null)
            val present = Holder("ok")
            __check((missing.value == null).toString(), "true")
            __check((present.value).toString(), "ok")
        }
