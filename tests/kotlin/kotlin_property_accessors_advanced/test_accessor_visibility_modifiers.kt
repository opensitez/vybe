// vybe-test: kotlin/kotlin_property_accessors_advanced/test_accessor_visibility_modifiers
// origin: languages/kotlin/tests/kotlin/test_kotlin_property_accessors_advanced.rs

class Item {
            var value: Int = 1
                private set
                public get
            fun add(v: Int) { value += v }
        }

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val i = Item()
            i.add(2)
            __check((i.value).toString(), "3")
        }
