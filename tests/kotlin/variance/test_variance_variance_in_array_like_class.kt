// vybe-test: kotlin/variance/test_variance_variance_in_array_like_class
// origin: languages/kotlin/tests/kotlin/test_variance.rs

class Box<out T>(val value: T)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val textBox: Box<String> = Box("v")
            val item: Box<Any> = textBox
            __check((item.value).toString(), "v")
        }
