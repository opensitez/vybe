use super::helpers::{compile_ok, run_ruby_one};

// -- Symbol#to_s converts to string

#[test]
fn symbol_to_s_converts_to_string() {
    compile_ok(
        "s = :hello.to_s
",
    );
}

#[test]
fn symbol_to_s_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :world.to_s
"
        ),
        "world"
    );
}

// -- Symbol#to_proc used with map

#[test]
fn symbol_to_proc_with_map() {
    compile_ok(
        "result = ['hello', 'world'].map(&:upcase)
",
    );
}

// -- Symbol#inspect wraps in colon

#[test]
fn symbol_inspect_wraps_in_colon() {
    compile_ok(
        "s = :hello.inspect
",
    );
}

#[test]
fn symbol_inspect_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :hello.inspect
"
        ),
        ":hello"
    );
}

// -- Symbol#id2name alias of to_s

#[test]
fn symbol_id2name_alias() {
    compile_ok(
        "s = :hello.id2name
",
    );
}

#[test]
fn symbol_id2name_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :world.id2name
"
        ),
        "world"
    );
}

// -- Symbol#length / #size character count

#[test]
fn symbol_length_character_count() {
    compile_ok(
        "n = :hello.length
",
    );
}

#[test]
fn symbol_size_alias() {
    compile_ok(
        "n = :hello.size
",
    );
}

#[test]
fn symbol_length_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :hello.length
"
        ),
        "5"
    );
}

// -- Symbol#upcase converts case

#[test]
fn symbol_upcase_converts_case() {
    compile_ok(
        "s = :hello.upcase
",
    );
}

#[test]
fn symbol_upcase_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :hello.upcase
"
        ),
        "HELLO"
    );
}

// -- Symbol#downcase converts case

#[test]
fn symbol_downcase_converts_case() {
    compile_ok(
        "s = :HELLO.downcase
",
    );
}

#[test]
fn symbol_downcase_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :HELLO.downcase
"
        ),
        "hello"
    );
}

// -- Symbol comparison with ==

#[test]
fn symbol_equality_comparison() {
    compile_ok(
        "result = :hello == :hello
",
    );
}

#[test]
fn symbol_equality_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts :foo == :foo
"
        ),
        "true"
    );
}

// -- Symbol <=> spaceship operator

#[test]
fn symbol_spaceship_operator() {
    compile_ok(
        "result = :apple <=> :banana
",
    );
}

#[test]
fn symbol_spaceship_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts (:apple <=> :banana)
"
        ),
        "-1"
    );
}

// -- Symbols are identical objects (same symbol === same object_id)

#[test]
fn symbol_identity_same_object_id() {
    compile_ok(
        "result = :hello.object_id == :hello.object_id
",
    );
}

// -- Symbol#to_proc in select

#[test]
fn symbol_to_proc_in_select() {
    compile_ok(
        "result = [1, 2, 3, nil, false].select(&:itself)
",
    );
}

// -- Symbol#to_proc in reject

#[test]
fn symbol_to_proc_in_reject() {
    compile_ok(
        "result = ['a', '', 'b', ''].reject(&:empty?)
",
    );
}

// -- Symbol#to_proc in sort_by

#[test]
fn symbol_to_proc_in_sort_by() {
    compile_ok(
        "words = ['banana', 'apple', 'cherry']
result = words.sort_by(&:length)
",
    );
}

// -- %i[] symbol array literal

#[test]
fn symbol_percent_i_literal() {
    compile_ok(
        "syms = %i[foo bar baz]
",
    );
}

#[test]
fn symbol_percent_i_runtime_length() {
    assert_eq!(
        run_ruby_one(
            "puts %i[foo bar baz].length
"
        ),
        "3"
    );
}

// -- Symbol as method call via send

#[test]
fn symbol_as_send_method_name() {
    compile_ok(
        "result = 'hello'.send(:upcase)
",
    );
}

#[test]
fn symbol_send_runtime() {
    assert_eq!(
        run_ruby_one(
            "puts 'hello'.send(:reverse)
"
        ),
        "olleh"
    );
}

// -- Symbol#match regex match

#[test]
fn symbol_match_regex() {
    compile_ok(
        "result = :hello.match(/ell/)
",
    );
}

// -- Symbol#encoding

#[test]
fn symbol_encoding() {
    compile_ok(
        "enc = :hello.encoding
",
    );
}

// -- Symbol.all_symbols (compile_ok)

#[test]
fn symbol_all_symbols() {
    compile_ok(
        "syms = Symbol.all_symbols
",
    );
}

// -- Symbol in case/when comparison

#[test]
fn symbol_in_case_when() {
    compile_ok(
        "status = :ok
result = case status
when :ok then 'good'
when :error then 'bad'
else 'unknown'
end
",
    );
}

#[test]
fn symbol_case_when_runtime() {
    assert_eq!(
        run_ruby_one(
            "status = :ok
result = case status
when :ok then 'good'
else 'other'
end
puts result
"
        ),
        "good"
    );
}

// -- Symbol hash key vs string hash key (different keys)

#[test]
fn symbol_hash_key_vs_string_key() {
    compile_ok(
        "h = {}
h[:foo] = 1
h['foo'] = 2
result = h[:foo] == h['foo']
",
    );
}

#[test]
fn symbol_hash_key_different_from_string_runtime() {
    assert_eq!(
        run_ruby_one(
            "h = {}
h[:foo] = 1
h['foo'] = 2
puts h.length
"
        ),
        "2"
    );
}
