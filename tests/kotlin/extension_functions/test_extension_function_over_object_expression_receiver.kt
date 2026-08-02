// vybe-test: kotlin/extension_functions/test_extension_function_over_object_expression_receiver
// origin: languages/kotlin/tests/kotlin/test_extension_functions.rs

fun StringBuilder.enclosed(): String {
            this.append("]")
            this.insert(0, "[")
            return this.toString()
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value = StringBuilder("ok").enclosed()
            __check((value).toString(), "[ok]")
        }
