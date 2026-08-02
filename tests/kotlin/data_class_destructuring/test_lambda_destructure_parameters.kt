// vybe-test: kotlin/data_class_destructuring/test_lambda_destructure_parameters
// origin: languages/kotlin/tests/kotlin/test_data_class_destructuring.rs

data class Item(val id: Int, val label: String)

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val values = listOf(Item(1, "a"), Item(2, "b"))
            val out = values.joinToString("-") { (id, label) -> "$id:$label" }
            __check((out).toString(), "1:a-2:b")
        }
