// vybe-test: kotlin/visibility/test_internal_member_allows_rebinding_within_same_module
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Box {
            internal var value: Int = 2
            fun bump() {
                value++
            }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val box = Box()
            box.bump()
            box.value = 9
            __check((box.value).toString(), "9")
        }
