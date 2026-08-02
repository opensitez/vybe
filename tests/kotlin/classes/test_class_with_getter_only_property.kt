// vybe-test: kotlin/classes/test_class_with_getter_only_property
// origin: languages/kotlin/tests/kotlin/test_classes.rs

class Product { val price = 7
val doubled: Int get() = price * 2 }
fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() { val p = Product()
__check((p.doubled).toString(), "14") }
