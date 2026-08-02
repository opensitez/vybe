// vybe-test: kotlin/constructor_chaining/test_constructor_mismatched_types_compile_path
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Num(val text: String) {
            constructor(v: Int) : this(v.toString())
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Num(10).text).toString(), "10")
        }
