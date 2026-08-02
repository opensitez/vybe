// vybe-test: kotlin/properties/test_property_extension_on_instance_has_read_access_only
// origin: languages/kotlin/tests/kotlin/test_properties.rs

class Item(val name: String)

        val Item.label: String
            get() = name + "-tag"

        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val item = Item("x")
            __check((item.label).toString(), "x-tag")
        }
