// vybe-test: kotlin/constructor_chaining/test_constructor_reified_like_no
// origin: languages/kotlin/tests/kotlin/test_constructor_chaining.rs

class Holder {
            val value: String
            constructor(v: Int) { value = v.toString() }
            constructor(v: String) { value = v }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Holder(4).value).toString(), "4")
            __check((Holder("x").value).toString(), "x")
        }
