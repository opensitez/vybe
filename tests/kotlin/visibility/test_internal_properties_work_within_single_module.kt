// vybe-test: kotlin/visibility/test_internal_properties_work_within_single_module
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            internal var tag: String = "mod"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            item.tag = "ok"
            __check((item.tag).toString(), "ok")
        }
