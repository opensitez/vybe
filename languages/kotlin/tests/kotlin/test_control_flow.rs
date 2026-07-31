use crate::helpers::run_prints;

#[test]
fn test_if_else() {
    let out = run_prints(r#"
        fun main() {
            val x = 10
            if (x > 5) {
                println("greater")
            } else {
                println("smaller")
            }
        }
    "#);
    assert_eq!(out, &["greater"]);
}

#[test]
fn test_if_expression() {
    let out = run_prints(r#"
        fun main() {
            val a = 7
            val b = 12
            val max = if (a > b) a else b
            println(max)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_while_loop() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            while (i < 3) {
                println(i)
                i += 1
            }
        }
    "#);
    assert_eq!(out, &["0", "1", "2"]);
}

#[test]
fn test_for_range() {
    let out = run_prints(r#"
        fun main() {
            for (i in 1..3) {
                println(i)
            }
        }
    "#);
    assert_eq!(out, &["1", "2", "3"]);
}

#[test]
fn test_do_while_loop() {
    let out = run_prints(r#"
        fun main() {
            var x = 5
            do {
                println(x)
                x -= 2
            } while (x > 0)
        }
    "#);
    assert_eq!(out, &["5", "3", "1"]);
}

#[test]
fn test_break_and_continue() {
    let out = run_prints(r#"
        fun main() {
            for (i in 1..5) {
                if (i == 2) continue
                if (i == 4) break
                println(i)
            }
        }
    "#);
    assert_eq!(out, &["1", "3"]);
}

#[test]
fn test_if_else_if_ladder() {
    let out = run_prints(r#"
        fun main() {
            val score = 85
            if (score >= 90) {
                println("A")
            } else if (score >= 80) {
                println("B")
            } else {
                println("C")
            }
        }
    "#);
    assert_eq!(out, &["B"]);
}

#[test]
fn test_nested_loops() {
    let out = run_prints(r#"
        fun main() {
            for (i in 1..2) {
                for (j in 1..2) {
                    println(i * 10 + j)
                }
            }
        }
    "#);
    assert_eq!(out, &["11", "12", "21", "22"]);
}

#[test]
fn test_while_with_break() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            while (true) {
                if (count == 3) break
                println(count)
                count += 1
            }
        }
    "#);
    assert_eq!(out, &["0", "1", "2"]);
}

#[test]
fn test_while_with_continue() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            while (i < 5) {
                i += 1
                if (i % 2 == 0) continue
                println(i)
            }
        }
    "#);
    assert_eq!(out, &["1", "3", "5"]);
}

#[test]
fn test_single_line_if() {
    let out = run_prints(r#"
        fun main() {
            val x = 10
            if (x > 0) println("positive")
        }
    "#);
    assert_eq!(out, &["positive"]);
}

#[test]
fn test_when_statement_basic() {
    let out = run_prints(r#"
        fun main() {
            val x = 2
            when (x) {
                1 -> println("one")
                2 -> println("two")
            }
        }
    "#);
    assert_eq!(out, &["two"]);
}

#[test]
fn test_when_statement_else() {
    let out = run_prints(r#"
        fun main() {
            val x = 99
            when (x) {
                1 -> println("one")
                else -> println("other")
            }
        }
    "#);
    assert_eq!(out, &["other"]);
}

#[test]
fn test_for_loop_accumulation() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 1..4) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_nested_if_expressions() {
    let out = run_prints(r#"
        fun main() {
            val a = 5
            val b = 10
            val c = 15
            val max = if (a > b) (if (a > c) a else c) else (if (b > c) b else c)
            println(max)
        }
    "#);
    assert_eq!(out, &["15"]);
}

#[test]
fn test_for_range_with_step() {
    let out = run_prints(r#"
        fun main() {
            var sum = 0
            for (i in 1..7 step 2) {
                sum += i
            }
            println(sum)
        }
    "#);
    assert_eq!(out, &["16"]);
}

#[test]
fn test_do_while_terminates() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            do {
                println(i)
                i += 1
            } while (i < 3)
        }
    "#);
    assert_eq!(out, &["0", "1", "2", "3"]);
}

#[test]
fn test_while_with_nested_break() {
    let out = run_prints(r#"
        fun main() {
            var outer = 0
            while (outer < 3) {
                var inner = 0
                while (inner < 3) {
                    if (inner == 1) break
                    println(outer * 10 + inner)
                    inner += 1
                }
                outer += 1
            }
        }
    "#);
    assert_eq!(out, &["0", "10", "20"]);
}

#[test]
fn test_when_as_expression_in_assignment() {
    let out = run_prints(r#"
        fun main() {
            val score = 77
            val grade = when (score) {
                in 90..100 -> "A"
                in 80..89 -> "B"
                in 70..79 -> "C"
                else -> "F"
            }
            println(grade)
        }
    "#);
    assert_eq!(out, &["C"]);
}

#[test]
fn test_if_guard_in_when() {
    let out = run_prints(r#"
        fun main() {
            val x = 10
            when {
                x < 0 -> println("negative")
                x == 0 -> println("zero")
                x > 0 -> println("positive")
            }
        }
    "#);
    assert_eq!(out, &["positive"]);
}

#[test]
fn test_if_else_nested_with_scopes() {
    let out = run_prints(r#"
        fun main() {
            val score = 82
            if (score >= 70) {
                if (score >= 80) {
                    println("pass-a")
                } else {
                    println("pass-b")
                }
            } else {
                println("fail")
            }
        }
    "#);
    assert_eq!(out, &["pass-a"]);
}

#[test]
fn test_while_not_entered() {
    let out = run_prints(r#"
        fun main() {
            var entered = false
            while (false) {
                entered = true
            }
            println(entered)
        }
    "#);
    assert_eq!(out, &["false"]);
}

#[test]
fn test_for_with_destructuring_loop() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for ((x, y) in arrayOf(Pair(1, 2), Pair(3, 4))) {
                total += x + y
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["10"]);
}

#[test]
fn test_do_while_runs_once_when_false() {
    let out = run_prints(r#"
        fun main() {
            var n = 0
            do {
                n += 1
            } while (n > 10)
            println(n)
        }
    "#);
    assert_eq!(out, &["1"]);
}

#[test]
fn test_while_with_continue_and_nested_if() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            var i = 0
            while (i < 6) {
                i += 1
                if (i == 2 || i == 5) {
                    continue
                }
                count += i
            }
            println(count)
        }
    "#);
    assert_eq!(out, &["12"]);
}

#[test]
fn test_for_range_down_to_accumulator() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 8 downTo 3) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["33"]);
}

#[test]
fn test_while_with_if_assignment() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            var i = 0
            while (i < 4) {
                val next = if (i % 2 == 0) i else i + 1
                total += next
                i += 1
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["8"]);
}

#[test]
fn test_while_inner_nested_break() {
    let out = run_prints(r#"
        fun main() {
            var outer = 0
            while (outer < 4) {
                var inner = 0
                while (inner < 4) {
                    inner += 1
                    if (inner == 3) {
                        break
                    }
                }
                outer += 1
                println(outer)
                if (outer == 2) {
                    break
                }
            }
        }
    "#);
    assert_eq!(out, &["1", "2"]);
}

#[test]
fn test_when_with_boolean_branches() {
    let out = run_prints(r#"
        fun main() {
            val isReady = true
            when {
                isReady == false -> println("no")
                isReady && true -> println("yes")
                else -> println("maybe")
            }
        }
    "#);
    assert_eq!(out, &["yes"]);
}

#[test]
fn test_if_expression_used_in_assignment() {
    let out = run_prints(r#"
        fun main() {
            val count = 6
            val label = if (count > 10) "large" else if (count >= 5) "medium" else "small"
            println(label)
        }
    "#);
    assert_eq!(out, &["medium"]);
}

#[test]
fn test_until_range_excludes_upper_bound() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (i in 1 until 4) {
                total += i
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_for_loop_with_indices_and_stride() {
    let out = run_prints(r#"
        fun main() {
            val nums = arrayOf(2, 4, 6, 8, 10)
            var total = 0
            for (i in nums.indices step 2) {
                total += nums[i]
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["18"]);
}

#[test]
fn test_while_with_postcondition_and_early_exit() {
    let out = run_prints(r#"
        fun main() {
            var i = 1
            var total = 0
            while (i <= 10) {
                total += i
                if (total >= 6) break
                i += 1
            }
            println(total)
            println(i)
        }
    "#);
    assert_eq!(out, &["6", "3"]);
}

#[test]
fn test_labeled_break_exits_outer_loop_only() {
    let out = run_prints(r#"
        fun main() {
            var values = ""
            outer@ for (outer in 1..3) {
                for (inner in 1..3) {
                    if (outer == 2 && inner == 2) {
                        break@outer
                    }
                    values += "${outer}${inner};"
                }
            }
            println(values)
        }
    "#);
    assert_eq!(out, &["111;112;21;"]);
}

#[test]
fn test_labeled_continue_skips_to_next_outer_iteration() {
    let out = run_prints(r#"
        fun main() {
            var values = ""
            outer@ for (outer in 1..3) {
                inner@ for (inner in 1..3) {
                    if (inner == 2) continue@outer
                    values += "${outer}${inner};"
                }
            }
            println(values)
        }
    "#);
    assert_eq!(out, &["11;21;31;"]);
}

#[test]
fn test_while_loop_as_expression_in_assignment() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            val sum = run {
                var acc = 0
                while (i < 3) {
                    acc += i
                    i += 1
                }
                acc
            }
            println(sum)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_when_with_guard_and_fallthrough_logic() {
    let out = run_prints(r#"
        fun main() {
            val score = 92
            val label = when (score) {
                in 90..100 -> if (score % 2 == 0) "A" else "A+"
                in 80..89 -> "B"
                else -> "F"
            }
            println(label)
        }
    "#);
    assert_eq!(out, &["A"]);
}

#[test]
fn test_when_with_is_type_checks() {
    let out = run_prints(r#"
        fun main() {
            val value: Any = "kotlin"
            val category = when (value) {
                is Int -> "int"
                is String -> "string"
                is Boolean -> "bool"
                else -> "other"
            }
            println(category)
        }
    "#);
    assert_eq!(out, &["string"]);
}

#[test]
fn test_if_expression_chain_and_type_flow() {
    let out = run_prints(r#"
        fun main() {
            val x = 10
            val tier = if (x > 20) "high" else if (x > 5) "mid" else "low"
            if (x > 5) {
                println(tier)
                println("statement")
            }
        }
    "#);
    assert_eq!(out, &["mid", "statement"]);
}

#[test]
fn test_do_while_with_continue_and_break() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            var out = ""
            do {
                i++
                if (i == 2) continue
                if (i > 4) break
                out += i.toString()
            } while (i < 6)
            println(out)
        }
    "#);
    assert_eq!(out, &["134"]);
}

#[test]
fn test_repeat_iterates_exact_times() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            repeat(4) { index ->
                total += index
            }
            println(total)
            var markers = ""
            repeat(0) {
                markers += "x"
            }
            println(markers.isEmpty())
        }
    "#);
    assert_eq!(out, &["6", "true"]);
}

#[test]
fn test_char_range_for_loop() {
    let out = run_prints(r#"
        fun main() {
            var letters = ""
            for (c in 'a'..'d') {
                letters += c
            }
            println(letters)
        }
    "#);
    assert_eq!(out, &["abcd"]);
}

#[test]
fn test_for_down_to_step_negative_offset() {
    let out = run_prints(r#"
        fun main() {
            var nums = ""
            for (n in 9 downTo 3 step 2) {
                nums += n.toString()
            }
            println(nums)
        }
    "#);
    assert_eq!(out, &["9753"]);
}

#[test]
fn test_while_loop_with_labeled_break() {
    let out = run_prints(r#"
        fun main() {
            var trace = ""
            var outer = 0
            outer@ while (outer < 4) {
                var inner = 0
                while (inner < 4) {
                    if (inner == 2) {
                        break@outer
                    }
                    trace += "${outer}-${inner};"
                    inner += 1
                }
                outer += 1
            }
            println(trace)
        }
    "#);
    assert_eq!(out, &["0-0;0-1;"]);
}

#[test]
fn test_while_loop_with_labeled_continue() {
    // The outer loop must be BOUNDED. `continue@outer` re-enters the outer
    // body, which re-runs `var inner = 0`, so `inner` never gets past 2 and a
    // trailing `break` is unreachable — the previous version of this test was
    // an infinite loop in real Kotlin too, and its expected "134" was the
    // answer for an UNLABELED `continue` (which stays in the inner loop).
    //
    // "111" is the discriminating answer: each round records `1`, then
    // `continue@outer` abandons the rest of the INNER loop *and* the rest of
    // the outer body. If the label were ignored this would print "134134134".
    let out = run_prints(r#"
        fun main() {
            var trace = ""
            var rounds = 0
            outer@ while (rounds < 3) {
                rounds += 1
                var inner = 0
                while (inner < 4) {
                    inner += 1
                    if (inner == 2) {
                        continue@outer
                    }
                    trace += inner.toString()
                }
                trace += "x"
            }
            println(trace)
        }
    "#);
    assert_eq!(out, &["111"]);
}

#[test]
fn test_while_else_if_in_loop_condition() {
    let out = run_prints(r#"
        fun main() {
            var i = 0
            var evenCount = 0
            var oddCount = 0
            while (i < 6) {
                if (i % 2 == 0) {
                    evenCount += 1
                } else {
                    oddCount += 1
                }
                i += 1
            }
            println(evenCount)
            println(oddCount)
        }
    "#);
    assert_eq!(out, &["3", "3"]);
}

#[test]
fn test_when_without_subject_uses_guard_chain() {
    let out = run_prints(r#"
        fun main() {
            val score = 58
            val band = when {
                score >= 90 -> "A"
                score >= 80 -> "B"
                score >= 70 -> "C"
                score >= 60 -> "D"
                else -> "F"
            }
            println(band)
        }
    "#);
    assert_eq!(out, &["D"]);
}

#[test]
fn test_if_expression_skips_false_branch_side_effects() {
    let out = run_prints(r#"
        var hits = 0

        fun bump(): Int {
            hits += 1
            return 0
        }

        fun main() {
            val value = if (1 == 1) {
                7
            } else {
                bump()
            }
            println(value)
            println(hits)
        }
    "#);
    assert_eq!(out, &["7", "0"]);
}

#[test]
fn test_for_loop_with_if_filter_in_body() {
    let out = run_prints(r#"
        fun main() {
            var values = ""
            for (i in 1..8) {
                if (i % 2 == 1) {
                    continue
                }
                values += i.toString()
            }
            println(values)
        }
    "#);
    assert_eq!(out, &["2468"]);
}

#[test]
fn test_while_loop_condition_function_calls_evaluate_per_iteration() {
    let out = run_prints(r#"
        var calls = 0

        fun shouldContinue(): Boolean {
            calls += 1
            return calls < 3
        }

        fun main() {
            var count = 0
            while (shouldContinue()) {
                count += 1
            }
            println(count)
            println(calls)
        }
    "#);
    assert_eq!(out, &["2", "3"]);
}

#[test]
fn test_labeled_continue_from_nested_for() {
    let out = run_prints(r#"
    fun main() {
            var out = ""
            outer@ for (row in 1..4) {
                for (col in 1..4) {
                    if (col == 3) continue@outer
                    out += "${row}${col}|"
                }
            }
            println(out)
        }
    "#);
    assert_eq!(out, &["11|12|21|22|31|32|41|42|"]);
}

#[test]
fn test_return_from_run_block_with_label() {
    let out = run_prints(r#"
        fun main() {
            val found = run {
                for (i in 1..4) {
                    if (i == 3) {
                        return@run i
                    }
                }
                0
            }
            println(found)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_when_as_statement_with_multiple_checks_same_line() {
    let out = run_prints(r#"
        fun main() {
            val x = 5
            when (x) {
                in 1..3, in 7..9 -> println("edge")
                5, 6 -> println("middle")
                else -> println("other")
            }
        }
    "#);
    assert_eq!(out, &["middle"]);
}

#[test]
fn test_nested_when_and_for_interaction() {
    let out = run_prints(r#"
        fun main() {
            val input = arrayOf(1, 2, 3, 4)
            var marks = ""
            for (value in input) {
                when {
                    value == 1 -> marks += "A"
                    value in 2..3 -> marks += "B"
                    else -> marks += "C"
                }
            }
            println(marks)
        }
    "#);
    assert_eq!(out, &["ABBC"]);
}

#[test]
fn test_do_while_condition_uses_updated_variable() {
    let out = run_prints(r#"
        fun main() {
            var x = 0
            do {
                println(x)
                x += 3
            } while (x < 8)
        }
    "#);
    assert_eq!(out, &["0", "3", "6"]);
}

#[test]
fn test_repeat_zero_with_side_effect() {
    let out = run_prints(r#"
        var seen = 0
        fun main() {
            repeat(0) {
                seen += 1
            }
            println(seen)
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_while_true_with_labeled_break_and_finally() {
    let out = run_prints(r#"
        fun main() {
            var steps = 0
            try {
                outer@ while (true) {
                    steps += 1
                    if (steps == 2) {
                        break@outer
                    }
                }
            } finally {
                println(steps)
            }
            println("done")
        }
    "#);
    assert_eq!(out, &["2", "done"]);
}

#[test]
fn test_repeat_negative_count_throws() {
    let out = run_prints(r#"
        fun main() {
            try {
                repeat(-1) {
                    println("bad")
                }
            } catch (e: Exception) {
                println("caught")
            }
        }
    "#);
    assert_eq!(out, &["caught"]);
}

#[test]
fn test_for_empty_down_to_range_skips_body() {
    let out = run_prints(r#"
        fun main() {
            var seen = ""
            for (i in 5 downTo 10) {
                seen += i.toString()
            }
            println(seen.isEmpty())
        }
    "#);
    assert_eq!(out, &["true"]);
}

#[test]
fn test_for_while_do_when_combination() {
    let out = run_prints(r#"
        fun main() {
            var trace = ""
            for (i in 1..6) {
                when (i % 3) {
                    0 -> continue
                    1 -> trace += "a"
                    else -> trace += "b"
                }
                if (i > 4) break
            }
            println(trace)
        }
    "#);
    assert_eq!(out, &["abab"]);
}

#[test]
fn test_repeat_zero_iteration() {
    let out = run_prints(r#"
        fun main() {
            var count = 0
            repeat(0) { count += 1 }
            println(count)
        }
    "#);
    assert_eq!(out, &["0"]);
}

#[test]
fn test_while_loop_condition_reflects_external_change() {
    let out = run_prints(r#"
        var done = false

        fun shouldRun(): Boolean {
            return !done
        }

        fun main() {
            var total = 0
            while (shouldRun()) {
                total += 1
                done = true
            }
            println(total)
            println(shouldRun())
        }
    "#);
    assert_eq!(out, &["1", "false"]);
}
