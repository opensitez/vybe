// vybe-test: kotlin/smart_casts/test_when_type_dispatch_with_three_branches
// origin: languages/kotlin/tests/kotlin/test_smart_casts.rs

class A
        class B : A()
        class C : A()

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: A = C()
            val label = when (value) {
                is B -> "b"
                is C -> "c"
                else -> "a"
            }
            __check((label).toString(), "c")
        }
