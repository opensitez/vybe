// vybe-test: kotlin/variance/test_variance_generic_function_out_type
// origin: languages/kotlin/tests/kotlin/test_variance.rs

interface Source<out T> {
            fun get(): T
        }
        class UserSource : Source<String> {
            override fun get(): String = "u"
        }
        fun wrap(source: Source<out Any>): Any {
            return source.get()
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((wrap(UserSource())).toString(), "u")
        }
