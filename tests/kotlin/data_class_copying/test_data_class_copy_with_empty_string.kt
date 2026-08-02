// vybe-test: kotlin/data_class_copying/test_data_class_copy_with_empty_string
// origin: languages/kotlin/tests/kotlin/test_data_class_copying.rs

data class Label(val text: String)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val a = Label("ok")
            val b = a.copy(text = "")
            __check((a.text.isNotEmpty()).toString(), "true")
            __check((b.text.isEmpty()).toString(), "true")
        }
