// vybe-test: kotlin/string_templates/test_template_custom_to_string
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

data class Point(val x: Int, val y: Int)
        fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val p = Point(1, 2)
            __check(("point=$p").toString(), "point=Point(x=1, y=2)")
        }
