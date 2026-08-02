// vybe-test: kotlin/visibility/test_visibility_default_is_public_for_top_level_members
// origin: languages/kotlin/tests/kotlin/test_visibility.rs

class Item {
            fun status(): String = "public"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item()
            __check((item.status()).toString(), "public")
        }
