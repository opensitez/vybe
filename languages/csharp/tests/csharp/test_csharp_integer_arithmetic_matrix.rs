use super::helpers::run_csharp;

fn bool_text(v: bool) -> &'static str {
    if v { "True" } else { "False" }
}

/// A C# `int[]` literal (`{-12, -7, ...}`) from a slice of values.
fn arr(values: &[i64]) -> String {
    let items: Vec<String> = values.iter().map(|v| v.to_string()).collect();
    format!("{{{}}}", items.join(", "))
}

// These matrices exercise integer arithmetic over every operand pair. The
// operand pairs are looped IN the compiled program (one compile, the matrix
// runs in WASM) rather than recompiling a fresh program per pair — the
// pipeline is the same for every pair, so 121+ full compiles just to vary two
// literals was pure overhead. The Rust side rebuilds the expected output in
// the same print order.

#[test]
fn int_matrix_add_sub_mul() {
    let values: [i64; 11] = [-12, -7, -3, -1, 0, 1, 2, 3, 4, 5, 8];
    let src = format!(
        "int[] vals = {vals};\n\
         foreach (int a in vals) foreach (int b in vals) {{ \
         Console.WriteLine(a + b); Console.WriteLine(a - b); Console.WriteLine(a * b); }}",
        vals = arr(&values)
    );
    let mut expected = Vec::new();
    for a in values {
        for b in values {
            expected.push((a + b).to_string());
            expected.push((a - b).to_string());
            expected.push((a * b).to_string());
        }
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn int_matrix_div_and_mod_with_safe_denominators() {
    let numerators: [i64; 11] = [-25, -10, -3, -1, 1, 2, 4, 7, 9, 12, 20];
    let denominators: [i64; 10] = [-10, -3, -1, 1, 2, 3, 4, 5, 6, 7];
    let src = format!(
        "int[] nums = {nums};\nint[] dens = {dens};\n\
         foreach (int a in nums) foreach (int b in dens) {{ \
         Console.WriteLine(a / b); Console.WriteLine(a % b); }}",
        nums = arr(&numerators),
        dens = arr(&denominators)
    );
    let mut expected = Vec::new();
    for a in numerators {
        for b in denominators {
            expected.push((a / b).to_string());
            expected.push((a % b).to_string());
        }
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn int_matrix_bitwise_ops() {
    let values: [i64; 11] = [-16, -9, -4, -1, 0, 1, 2, 4, 7, 10, 31];
    let src = format!(
        "int[] vals = {vals};\n\
         foreach (int a in vals) foreach (int b in vals) {{ \
         Console.WriteLine(a & b); Console.WriteLine(a | b); Console.WriteLine(a ^ b); }}",
        vals = arr(&values)
    );
    let mut expected = Vec::new();
    for a in values {
        for b in values {
            expected.push((a & b).to_string());
            expected.push((a | b).to_string());
            expected.push((a ^ b).to_string());
        }
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn int_matrix_shift_relations_and_unary_ops() {
    let values: [i64; 11] = [-16, -9, -4, -2, -1, 0, 1, 2, 4, 8, 15];
    let shifts: [i64; 6] = [0, 1, 2, 3, 4, 5];
    let src = format!(
        "int[] vals = {vals};\nint[] shifts = {shifts};\n\
         foreach (int a in vals) foreach (int shift in shifts) {{ \
         Console.WriteLine(a << shift); Console.WriteLine(a >> shift); }}\n\
         foreach (int a in vals) {{ Console.WriteLine(-a); Console.WriteLine(+a); }}",
        vals = arr(&values),
        shifts = arr(&shifts)
    );
    let mut expected = Vec::new();
    for a in values {
        for shift in shifts {
            expected.push((a << shift).to_string());
            expected.push((a >> shift).to_string());
        }
    }
    for a in values {
        expected.push((-a).to_string());
        expected.push(a.to_string());
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn int_matrix_comparisons_and_bool_compositions() {
    let values: [i64; 11] = [-12, -7, -1, 0, 1, 2, 3, 5, 8, 9, 13];
    let src = format!(
        "int[] vals = {vals};\n\
         foreach (int a in vals) foreach (int b in vals) {{ \
         Console.WriteLine(a > b); Console.WriteLine(a >= b); Console.WriteLine(a < b); \
         Console.WriteLine(a == b); Console.WriteLine(a != b); }}\n\
         foreach (int a in vals) {{ \
         Console.WriteLine(a % 2 == 0); Console.WriteLine(a > 0 && a % 2 == 0); \
         bool result = (a >= 0 || a < -1) && (a != 2); Console.WriteLine(result); }}",
        vals = arr(&values)
    );
    let mut expected = Vec::new();
    for a in values {
        for b in values {
            expected.push(bool_text(a > b).to_string());
            expected.push(bool_text(a >= b).to_string());
            expected.push(bool_text(a < b).to_string());
            expected.push(bool_text(a == b).to_string());
            expected.push(bool_text(a != b).to_string());
        }
    }
    for a in values {
        expected.push(bool_text(a % 2 == 0).to_string());
        expected.push(bool_text(a > 0 && a % 2 == 0).to_string());
        expected.push(bool_text((a >= 0 || a < -1) && (a != 2)).to_string());
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn integer_to_float_and_back_truncation_matrix() {
    let values: [i64; 11] = [-13, -7, -1, 0, 1, 2, 3, 5, 11, 27, 64];
    let src = format!(
        "int[] vals = {vals};\n\
         foreach (int n in vals) {{ \
         double ratio = n * 1.25; Console.WriteLine(ratio); \
         double valueD = n + 0.75; Console.WriteLine((int)valueD); \
         long widened = (long)n; Console.WriteLine(widened); \
         long source = (long)n + 7; int back = (int)source; Console.WriteLine(back); }}",
        vals = arr(&values)
    );
    let mut expected = Vec::new();
    for a in values {
        expected.push(format!("{}", (a as f64) * 1.25));
        expected.push(((a as f64 + 0.75).trunc() as i64).to_string());
        expected.push((a as i64).to_string());
        expected.push((a + 7).to_string());
    }
    assert_eq!(run_csharp(&src), expected);
}

#[test]
fn int_matrix_nullish_arithmetic_guarded_branches() {
    let thresholds: [i64; 8] = [0, 1, 2, 3, 5, 8, 13, 21];
    let src = format!(
        "int[] th = {th};\n\
         foreach (int a in th) foreach (int b in th) {{ \
         int result = a == 0 ? b : a + b; Console.WriteLine(result); \
         int guarded = b == 0 ? 0 : a / b; Console.WriteLine(guarded); }}",
        th = arr(&thresholds)
    );
    let mut expected = Vec::new();
    for a in thresholds {
        for b in thresholds {
            expected.push(if a == 0 { b } else { a + b }.to_string());
            expected.push(if b == 0 { 0 } else { a / b }.to_string());
        }
    }
    assert_eq!(run_csharp(&src), expected);
}
