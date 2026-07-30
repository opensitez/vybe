use crate::helpers::run_prints;

#[test]
fn test_top_level_and_local_scope_isolation() {
    let out = run_prints(r#"
        val label = "global"

        fun echo(value: String): String {
            return label + ":" + value
        }

        fun main() {
            val label = "local"
            println(echo("one"))
            println(label)
        }
    "#);
    assert_eq!(out, &["global:one", "local"]);
}

#[test]
fn test_shadowing_in_nested_blocks() {
    let out = run_prints(r#"
        fun main() {
            val mode = "outer"
            println(mode)
            {
                val mode = "inner"
                println(mode)
            }
            println(mode)
        }
    "#);
    assert_eq!(out, &["outer", "inner", "outer"]);
}

#[test]
fn test_scope_in_function_parameter_binding() {
    let out = run_prints(r#"
        fun square(value: Int): Int {
            return value * value
        }

        fun main() {
            val value = 3
            fun emit(value: Int): Int {
                return value + 1
            }
            println(square(value))
            println(emit(4))
            println(value)
        }
    "#);
    assert_eq!(out, &["9", "5", "3"]);
}

#[test]
fn test_local_function_captures_enclosing_scope() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            fun add(step: Int) {
                total += step
            }
            add(3)
            add(4)
            println(total)
        }
    "#);
    assert_eq!(out, &["7"]);
}

#[test]
fn test_scope_limited_block_variable_lifetime() {
    let out = run_prints(r#"
        fun main() {
            var value = 1
            if (value == 1) {
                val block = 8
                value += block
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["9"]);
}

#[test]
fn test_member_and_receiver_scope_split() {
    let out = run_prints(r#"
        class ScopeProbe {
            val value = 2
            fun combine(): Int {
                val value = 4
                return this.value + value
            }
        }

        fun main() {
            println(ScopeProbe().combine())
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_nested_block_rebinds_local_name() {
    let out = run_prints(r#"
        fun main() {
            val token = "global"
            {
                val token = "inner"
                println(token)
            }
            println(token)
        }
    "#);
    assert_eq!(out, &["inner", "global"]);
}

#[test]
fn test_function_parameter_shadows_outer_value() {
    let out = run_prints(r#"
        fun compute(value: Int): Int {
            return value * 2
        }

        fun main() {
            val value = 5
            println(compute(value))
            println(value)
        }
    "#);
    assert_eq!(out, &["10", "5"]);
}

#[test]
fn test_local_function_reads_enclosing_immutables() {
    let out = run_prints(r#"
        fun main() {
            val base = 10
            fun bump(): Int {
                return base + 1
            }
            fun bumpTwice(): Int {
                return bump() + 1
            }
            println(bump())
            println(bumpTwice())
        }
    "#);
    assert_eq!(out, &["11", "12"]);
}

#[test]
fn test_local_function_shadow_parameter() {
    let out = run_prints(r#"
        fun main() {
            val value = 3
            fun inner(value: Int): Int {
                return value + 1
            }
            println(inner(4))
            println(value)
        }
    "#);
    assert_eq!(out, &["5", "3"]);
}

#[test]
fn test_lambda_reads_outer_var_before_and_after_change() {
    let out = run_prints(r#"
        fun main() {
            var total = 1
            val addOne = { total += 1 }
            addOne()
            println(total)
            total = 10
            addOne()
            println(total)
        }
    "#);
    assert_eq!(out, &["2", "11"]);
}

#[test]
fn test_lambda_parameter_shadowing_outer_name() {
    let out = run_prints(r#"
        fun main() {
            var value = 7
            val toString = { value: Int -> value + 1 }
            println(value)
            println(toString(3))
            println(value)
        }
    "#);
    assert_eq!(out, &["7", "4", "7"]);
}

#[test]
fn test_if_else_scoped_binding() {
    let out = run_prints(r#"
        fun pick(flag: Boolean): String {
            return if (flag) {
                val value = "yes"
                value
            } else {
                val value = "no"
                value
            }
        }

        fun main() {
            println(pick(true))
            println(pick(false))
        }
    "#);
    assert_eq!(out, &["yes", "no"]);
}

#[test]
fn test_when_branch_scoped_names() {
    let out = run_prints(r#"
        fun describe(value: Int): String {
            return when (value) {
                1 -> {
                    val label = "one"
                    label
                }
                2 -> {
                    val label = "two"
                    label
                }
                else -> {
                    val label = "other"
                    label
                }
            }
        }

        fun main() {
            println(describe(1))
            println(describe(2))
            println(describe(4))
        }
    "#);
    assert_eq!(out, &["one", "two", "other"]);
}

#[test]
fn test_loop_index_scope_with_outer_name() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            val label = 5
            for (label in arrayOf(1, 2, 3)) {
                total += label
            }
            println(label)
            println(total)
        }
    "#);
    assert_eq!(out, &["5", "6"]);
}

#[test]
fn test_while_scope_and_mutation() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            var index = 0
            while (index < 3) {
                val step = index + 1
                total += step
                index += 1
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["6"]);
}

#[test]
fn test_nested_scope_with_return_value_capture() {
    let out = run_prints(r#"
        fun make() : Int {
            val prefix = 1
            fun inner(): Int {
                val suffix = 2
                return prefix + suffix
            }
            return inner()
        }

        fun main() {
            println(make())
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_block_reused_name_after_scope() {
    let out = run_prints(r#"
        fun main() {
            var name = "outer"
            {
                val name = "inner"
                println(name)
            }
            name = "next"
            println(name)
        }
    "#);
    assert_eq!(out, &["inner", "next"]);
}

#[test]
fn test_try_catch_scope_separation() {
    let out = run_prints(r#"
        fun main() {
            val state = "ok"
            try {
                throw Exception("bad")
            } catch (e: Exception) {
                val state = "caught"
                println(state)
            }
            println(state)
        }
    "#);
    assert_eq!(out, &["caught", "ok"]);
}

#[test]
fn test_nested_try_without_escape() {
    let out = run_prints(r#"
        fun main() {
            val value = 1
            try {
                val value = "inner"
                println(value)
            } catch (e: Exception) {
                println("err")
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["inner", "1"]);
}

#[test]
fn test_catch_binding_does_not_clobber_outer_scope() {
    let out = run_prints(r#"
        fun main() {
            val message = "root"
            try {
                throw Exception("boom")
            } catch (message: Exception) {
                println(message.message)
            }
            println("root")
        }
    "#);
    assert_eq!(out, &["boom", "root"]);
}

#[test]
fn test_scope_split_between_fields_and_locals() {
    let out = run_prints(r#"
        class Probe {
            val source = "field"
            fun valueOf(input: String): String {
                val source = input
                return this.source + "-" + source
            }
        }

        fun main() {
            println(Probe().valueOf("local"))
        }
    "#);
    assert_eq!(out, &["field-local"]);
}

#[test]
fn test_scope_in_function_inside_object_expression() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            val maker = object : Any() {
                fun add(value: Int): Int {
                    fun inner(): Int {
                        return value + 1
                    }
                    return inner()
                }
            }
            total += maker.add(2)
            println(total)
        }
    "#);
    assert_eq!(out, &["3"]);
}

#[test]
fn test_local_scope_inside_class_member() {
    let out = run_prints(r#"
        class Box {
            var value = 0
            fun addStep(step: Int): Int {
                val value = step
                fun bump(): Int {
                    return this.value + value
                }
                this.value = step * 2
                return bump()
            }
        }

        fun main() {
            val b = Box()
            println(b.addStep(4))
            println(b.value)
        }
    "#);
    assert_eq!(out, &["4", "8"]);
}

#[test]
fn test_nested_blocks_inside_expression() {
    let out = run_prints(r#"
        fun main() {
            val value = 2
            val result = {
                val value = 5
                value * 3
            }
            println(value)
            println(result)
        }
    "#);
    assert_eq!(out, &["2", "15"]);
}

#[test]
fn test_scope_after_if_scope_restore() {
    let out = run_prints(r#"
        fun main() {
            val value = "start"
            if (true) {
                val value = "if-branch"
                println(value)
            }
            println(value)
        }
    "#);
    assert_eq!(out, &["if-branch", "start"]);
}

#[test]
fn test_scope_in_nested_for_loops() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            for (row in arrayOf(arrayOf(1, 2), arrayOf(3, 4))) {
                for (col in row) {
                    val value = col * 2
                    total += value
                }
            }
            println(total)
        }
    "#);
    assert_eq!(out, &["20"]);
}

#[test]
fn test_scope_with_shadowed_loop_variables() {
    let out = run_prints(r#"
        fun main() {
            val i = 100
            var output = 0
            for (i in arrayOf(1, 2, 3)) {
                output += i
            }
            println(i)
            println(output)
        }
    "#);
    assert_eq!(out, &["100", "6"]);
}

#[test]
fn test_outer_scope_after_lambda_call() {
    let out = run_prints(r#"
        fun main() {
            val value = "outer"
            val action = {
                val value = "inner"
                println(value)
            }
            action()
            println(value)
        }
    "#);
    assert_eq!(out, &["inner", "outer"]);
}

#[test]
fn test_nested_local_fun_uses_outer_var_and_updates_var() {
    let out = run_prints(r#"
        fun main() {
            var value = 1
            fun inc(step: Int) {
                value += step
            }
            inc(2)
            val value = 9
            println(value)
            inc(3)
            println(value)
            println(value == 9)
        }
    "#);
    assert_eq!(out, &["9", "9", "true"]);
}

#[test]
fn test_inner_scope_does_not_modify_shadowed_outer() {
    let out = run_prints(r#"
        fun main() {
            val marker = "root"
            if (marker == "root") {
                val marker = "inner"
                println(marker)
            }
            println(marker)
        }
    "#);
    assert_eq!(out, &["inner", "root"]);
}

#[test]
fn test_scope_across_returned_function_body() {
    let out = run_prints(r#"
        fun makeGreeter(prefix: String): (Int) -> String {
            val suffix = "!"
            return { value ->
                val body = prefix + value.toString()
                body + suffix
            }
        }

        fun main() {
            val greet = makeGreeter("x")
            println(greet(1))
            println(greet(2))
        }
    "#);
    assert_eq!(out, &["x1!", "x2!"]);
}

#[test]
fn test_function_literal_scope_with_block_shadowing() {
    let out = run_prints(r#"
        fun main() {
            val value = 1
            val add = { input: Int ->
                val result = input + 1
                result
            }
            println(add(3))
            println(value)
        }
    "#);
    assert_eq!(out, &["4", "1"]);
}

#[test]
fn test_scope_after_smart_cast_branch_isolated_by_type() {
    let out = run_prints(r#"
        fun label(value: Any): String {
            return when (value) {
                is String -> "str:" + value.length
                is Int -> "int:" + value
                is Boolean -> "bool:" + value
                else -> "other"
            }
        }

        fun main() {
            println(label("abc"))
            println(label(9))
            println(label(true))
            println(label(2.5))
        }
    "#);
    assert_eq!(out, &["str:3", "int:9", "bool:true", "other"]);
}

#[test]
fn test_scope_nested_function_mutates_outer_var_after_shadowing() {
    let out = run_prints(r#"
        fun main() {
            var value = 5

            fun bump(delta: Int) {
                fun total(): Int {
                    return value + delta
                }
                value = total()
            }

            bump(3)
            println(value)

            val value = 10
            fun useShadowed(): Int {
                return value + 1
            }
            println(useShadowed())
        }
    "#);
    assert_eq!(out, &["8", "11"]);
}

#[test]
fn test_scope_in_for_each_lambda_and_outer_capture() {
    let out = run_prints(r#"
        fun main() {
            var total = 0
            val source = listOf(1, 2, 3)
            source.forEach {
                val transformed = it * 2
                total += transformed
            }
            println(total)
            println(source.size)
        }
    "#);
    assert_eq!(out, &["12", "3"]);
}

#[test]
fn test_scope_qualified_this_preserves_outer_reference() {
    let out = run_prints(r#"
        class Container {
            val factor = 3

            fun makeTag(input: String): String {
                return with(this) {
                    val factor = 10
                    input + "-" + this@Container.factor
                }
            }
        }

        fun main() {
            val c = Container()
            println(c.makeTag("id"))
        }
    "#);
    assert_eq!(out, &["id-3"]);
}

#[test]
fn test_scope_object_expression_captures_outer_binding() {
    let out = run_prints(r#"
        open class Base(val label: String)

        fun main() {
            var prefix = "one"
            val instance = object : Base("base") {
                val captured = prefix
            }
            println(instance.captured)
            prefix = "two"
            println(instance.label)
        }
    "#);
    assert_eq!(out, &["one", "base"]);
}

#[test]
fn test_scope_try_catch_variable_not_visible_after_block() {
    let out = run_prints(r#"
        fun main() {
            val status = "start"
            val result = try {
                throw Exception("boom")
            } catch (failure: Exception) {
                val status = "caught"
                status
            } finally {
                println(status)
            }
            println(result)
        }
    "#);
    assert_eq!(out, &["start", "caught"]);
}

#[test]
fn test_scope_receiver_and_outer_this_with_with_expression() {
    let out = run_prints(r#"
        class Holder {
            val source = "outer"
            fun label(tag: String): String {
                return with(this) {
                    val source = tag
                    source + "-" + this@Holder.source
                }
            }
        }

        fun main() {
            println(Holder().label("inner"))
        }
    "#);
    assert_eq!(out, &["inner-outer"]);
}

#[test]
fn test_scope_destructuring_bindings_are_block_local() {
    let out = run_prints(r#"
        fun main() {
            val (first, second) = Pair("left", "right")

            val inner = run {
                val (first, second) = Pair("inner-left", "inner-right")
                first + ":" + second
            }

            println(first)
            println(second)
            println(inner)
        }
    "#);
    assert_eq!(out, &["left", "right", "inner-left:inner-right"]);
}

#[test]
fn test_scope_if_expression_has_its_own_binding_scope() {
    let out = run_prints(r#"
        fun main() {
            val label = "outer"
            val result = if (true) {
                val label = "then"
                label
            } else {
                val label = "else"
                label
            }

            println(label)
            println(result)
        }
    "#);
    assert_eq!(out, &["outer", "then"]);
}

#[test]
fn test_scope_lambda_sees_updated_outer_binding() {
    let out = run_prints(r#"
        fun main() {
            var prefix = "before"
            val format = { value: String ->
                prefix + ":" + value
            }

            println(format("one"))

            prefix = "after"
            println(format("two"))
        }
    "#);
    assert_eq!(out, &["before:one", "after:two"]);
}

#[test]
fn test_scope_shadowed_loop_variable_stays_local_to_iteration_body() {
    let out = run_prints(r#"
        fun main() {
            val values = arrayOf(1, 2, 3)
            var sum = 0

            for (value in values) {
                run {
                    val value = value * 10
                    sum += value
                }
            }

            println(sum)
        }
    "#);
    assert_eq!(out, &["60"]);
}

#[test]
fn test_scope_function_apply_returns_receiver_reference() {
    let out = run_prints(r#"
        fun main() {
            val text = StringBuilder()
                .apply {
                    append("k")
                    append("otlin")
                }
            println(text.toString())
            println(text === StringBuilder("kotlin"))
        }
    "#);
    assert_eq!(out, &["kotlin", "false"]);
}

#[test]
fn test_scope_function_let_scopes_nullable_and_outer_state() {
    let out = run_prints(r#"
        fun main() {
            val prefix = "x"
            val result = prefix.let {
                val suffix = it.toUpperCase()
                suffix + "!"
            }
            println(result)
            println(prefix)
        }
    "#);
    assert_eq!(out, &["X!", "x"]);
}

#[test]
fn test_scope_function_also_chain_and_outer_capture() {
    let out = run_prints(r#"
        fun main() {
            var marker = "base"
            val values = mutableListOf(1, 2, 3)
                .also {
                    it.add(4)
                }
                .also {
                    marker = "after"
                }

            println(values.joinToString(","))
            println(marker)
        }
    "#);
    assert_eq!(out, &["1,2,3,4", "after"]);
}

#[test]
fn test_scope_function_with_receiver_preserves_outer_scope_name() {
    let out = run_prints(r#"
        class Box {
            val value = "outer"
            fun label(): String {
                return with(this) {
                    val value = "inner"
                    value + "-" + this.value
                }
            }
        }

        fun main() {
            println(Box().label())
        }
    "#);
    assert_eq!(out, &["inner-outer"]);
}

#[test]
fn test_scope_try_finally_alters_and_observes_outer_mutation() {
    let out = run_prints(r#"
        fun main() {
            var state = "open"
            try {
                state = "processing"
            } finally {
                state = state + "-done"
            }
            println(state)
        }
    "#);
    assert_eq!(out, &["processing-done"]);
}
