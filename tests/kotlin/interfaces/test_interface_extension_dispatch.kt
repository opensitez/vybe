// vybe-test: kotlin/interfaces/test_interface_extension_dispatch
// origin: languages/kotlin/tests/kotlin/test_interfaces.rs

interface Taggable {
            fun tag(): String
        }

        class Item : Taggable {
            override fun tag(): String = "item"
        }

        fun Taggable.label(): String {
            return this.tag() + "#"
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val value: Taggable = Item()
            __check((value.label()).toString(), "item#")
        }
