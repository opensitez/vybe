// vybe-test: kotlin/advanced_features/test_advanced_looped_while_for
// origin: languages/kotlin/tests/kotlin/test_advanced_features.rs

fun main() { var total = 0
for (i in 1..3) { total += i }
var x = 0
while (x < 2) { total += 1
x += 1 }
println(total) }

