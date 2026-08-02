// vybe-test: kotlin/visibility/test_public_members_are_implicitly_accessible
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            fun show(): String = "ok"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            __check((Item().show()).toString(), "ok")
        }
