// vybe-test: kotlin/member_references/test_bound_method_reference_with_local_value
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Holder(val tag: String) {
            fun emit() = tag
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val h = Holder("ok")
            val f = h::emit
            __check((f()).toString(), "ok")
        }
