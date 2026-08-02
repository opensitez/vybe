// vybe-test: kotlin/visibility/test_public_members_are_callable_by_default
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            val label = "x"
            fun text(): String = "ok"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            __check((item.label).toString(), "x")
            __check((item.text()).toString(), "ok")
        }
