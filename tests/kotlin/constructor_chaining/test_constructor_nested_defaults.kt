// vybe-test: kotlin/constructor_chaining/test_constructor_nested_defaults
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class C(val a: Int, val b: Int = 2) {
            constructor() : this(0)
            constructor(a: Int, b: Int, c: Int) : this(a + b + c)
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((C().b).toString(), "2")
            __check((C(1, 2).a).toString(), "1")
            __check((C(1, 2, 3).a).toString(), "6")
        }
