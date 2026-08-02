// vybe-test: kotlin/local_classes/test_local_interface_and_impl
// origin: languages/kotlin/tests/kotlin/test_local_classes.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            interface I { fun value(): Int }
            class C(val v: Int) : I { override fun value() = v }
            __check((C(4).value()).toString(), "4")
        }
