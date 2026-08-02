// vybe-test: kotlin/string_templates/test_template_with_nested_list
// origin: languages/kotlin/tests/kotlin/test_string_templates.rs

fun __check(got: String, want: String) {
    if (got != want) {
        println("FAIL: want [" + want + "] got [" + got + "]")
        throw Exception("assertion failed")
    }
}

fun main() {
            val grid = listOf(listOf(1, 2), listOf(3, 4))
            __check(("rows=${grid.size} first=${grid[0][0]}").toString(), "rows=2 first=1")
        }
