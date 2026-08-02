// vybe-test: kotlin/default_arguments/test_default_arguments_default_on_local_val_return
// origin: languages/kotlin/tests/kotlin/test_default_arguments.rs

fun make(label: String = "x", amount: Int = 3): String {
            return label + amount
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val base = make()
            val changed = make(label = "y", amount = 1)
            __check((base).toString(), "x3")
            __check((changed).toString(), "y1")
        }
