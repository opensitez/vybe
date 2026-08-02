// vybe-test: kotlin/classes/test_class_with_primary_default
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Item(val name: String = "a") { fun mainName(): String = name }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { __check((Item().mainName()).toString(), "a") }
