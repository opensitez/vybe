// vybe-test: kotlin/advanced_features/test_extension_function
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

class Text(val value: String)

        fun Text.emphasize(): String {
            return value + value
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val text = Text("hello")
            __check((text.emphasize()).toString(), "hellohello")
        }
