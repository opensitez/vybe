// vybe-test: kotlin/constructor_chaining/test_constructor_overload_count
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Over {
            val label: String
            constructor() { label = "x" }
            constructor(v: Int) { label = v.toString() }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Over().label).toString(), "x")
            __check((Over(2).label).toString(), "2")
        }
