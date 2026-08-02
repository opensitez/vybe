// vybe-test: kotlin/type_casts/test_safe_cast_to_missing_interface_is_null
// origin: languages/kotlin/tests/kotlin/test_type_casts.rs

interface Aware { fun marker(): String }
        interface Other { fun other(): Int }
        class Item : Aware { override fun marker(): String = "x" }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Any = Item()
            val casted = value as? Other
            __check((casted == null).toString(), "true")
        }
