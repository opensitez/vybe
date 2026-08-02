// vybe-test: kotlin/member_references/test_member_function_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Counter(val step: Int) {
            fun plus(v: Int) = v + step
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val add = Counter(5)::plus
            __check((add(7)).toString(), "12")
        }
