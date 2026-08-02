// vybe-test: kotlin/visibility/test_internal_property_access_within_module
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            internal val value = 9
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            __check((item.value).toString(), "9")
        }
