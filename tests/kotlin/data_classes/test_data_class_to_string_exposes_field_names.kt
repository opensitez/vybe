// vybe-test: kotlin/data_classes/test_data_class_to_string_exposes_field_names
// origin: languages/kotlin/tests/kotlin/test_data_classes.rs

data class Trace(val action: String, val count: Int)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Trace("run", 3)
            __check((item.toString()).toString(), "Trace(action=run, count=3)")
        }
