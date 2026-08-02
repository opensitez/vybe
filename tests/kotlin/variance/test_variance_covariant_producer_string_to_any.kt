// vybe-test: kotlin/variance/test_variance_covariant_producer_string_to_any
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Producer<out T> { fun provide(): T }
        class StringSource : Producer<String> {
            override fun provide(): String = "x"
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val source: Producer<Any> = StringSource()
            __check((source.provide()).toString(), "x")
        }
