use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Random & Secrets PRNG — random.seed, randint, choice, sample, shuffle, secrets.token_hex, choice, compare_digest
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_random_seed_reproducibility() {
    let src = r#"
import random

random.seed(42)
r1 = [random.randint(1, 100) for _ in range(5)]

random.seed(42)
r2 = [random.randint(1, 100) for _ in range(5)]

print(r1 == r2)
print(r1)
"#;
    assert_eq!(run_python(src), vec!["True", "[82, 15, 4, 95, 36]"]);
}

#[test]
fn test_py_random_choice_and_choices() {
    let src = r#"
import random

random.seed(123)
items = ["apple", "banana", "cherry", "date"]
chosen = random.choice(items)
print(chosen in items)

choices = random.choices(items, weights=[10, 1, 1, 1], k=5)
print(len(choices) == 5)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_random_sample_without_replacement() {
    let src = r#"
import random

random.seed(99)
population = range(1, 100)
sample = random.sample(population, 5)
print(len(sample))
print(len(set(sample)) == 5)  # all unique
"#;
    assert_eq!(run_python(src), vec!["5", "True"]);
}

#[test]
fn test_py_random_shuffle_in_place() {
    let src = r#"
import random

random.seed(77)
lst = [1, 2, 3, 4, 5]
random.shuffle(lst)
print(sorted(lst) == [1, 2, 3, 4, 5])
print(lst != [1, 2, 3, 4, 5])
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_random_distributions_uniform_gauss() {
    let src = r#"
import random

random.seed(10)
u = random.uniform(1.0, 5.0)
print(1.0 <= u <= 5.0)

g = random.gauss(0.0, 1.0)
print(isinstance(g, float))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_secrets_token_hex_bytes_urlsafe() {
    let src = r#"
import secrets

token = secrets.token_hex(16)
print(len(token))  # 32 hex chars for 16 bytes

raw_bytes = secrets.token_bytes(16)
print(len(raw_bytes))

url_token = secrets.token_urlsafe(16)
print(isinstance(url_token, str))
"#;
    assert_eq!(run_python(src), vec!["32", "16", "True"]);
}

#[test]
fn test_py_secrets_choice_secure() {
    let src = r#"
import secrets

chars = "abcdefghijklmnopqrstuvwxyz0123456789"
pwd = "".join(secrets.choice(chars) for _ in range(12))
print(len(pwd))
print(all(c in chars for c in pwd))
"#;
    assert_eq!(run_python(src), vec!["12", "True"]);
}

#[test]
fn test_py_secrets_compare_digest_constant_time() {
    let src = r#"
import secrets

h1 = "a" * 32
h2 = "a" * 32
h3 = "b" * 32

print(secrets.compare_digest(h1, h2))
print(secrets.compare_digest(h1, h3))
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_random_randrange_step() {
    let src = r#"
import random

random.seed(50)
even_val = random.randrange(0, 100, 2)
print(even_val % 2 == 0)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_random_system_random() {
    let src = r#"
import random

sr = random.SystemRandom()
val = sr.randint(1, 10)
print(1 <= val <= 10)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
