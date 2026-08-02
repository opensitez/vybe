// vybe-test: kotlin/member_references/test_constructor_reference
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Box(val value: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val ctor = ::Box
            val x = ctor(4)
            __check((x.value).toString(), "4")
        }
