// vybe-test: kotlin/scoping_functions/test_apply_and_run_builder_style
// origin: languages/kotlin/tests/kotlin/test_scoping_functions.rs

class Builder {
            var text = ""
            fun build(): String = text
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val result = Builder()
                .apply { text = "a" }
                .apply { text += "b" }
                .run { build() }
            __check((result).toString(), "ab")
        }
