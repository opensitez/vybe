// vybe-test: kotlin/member_references/test_reference_to_instance_of_nested_object
// origin: languages/kotlin/tests/kotlin/test_member_references.rs

class Box {
            inner class Inner {
                fun value(v: Int) = v + 1
            }
        }
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val f = Box().Inner()::value
            __check((f(8)).toString(), "9")
        }
