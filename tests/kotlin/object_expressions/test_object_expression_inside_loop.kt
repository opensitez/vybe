// vybe-test: kotlin/object_expressions/test_object_expression_inside_loop
// origin: languages/kotlin/tests/kotlin/test_object_expressions.rs

fun main() { val obj = object { var value = 0
fun inc() { value += 1 } }
for (i in 1..3) { obj.inc() }
println(obj.value) }

