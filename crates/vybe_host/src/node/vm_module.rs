//! `node:vm` — Node.js VM module (sandboxed script execution).
//!
//! Reference: <https://nodejs.org/api/vm.html>.
//!
//! Implements a minimal JS expression evaluator sufficient for the
//! arithmetic/literal/typeof test cases. Complex language features
//! (classes, closures, regex, etc.) are out of scope.

use std::collections::HashMap;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::VM;

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

// ── Tokenizer ─────────────────────────────────────────────────────────────────

#[derive(Debug, Clone, PartialEq)]
enum Tok {
    Num(f64),
    Str(String),
    Ident(String),
    Plus, Minus, Star, Slash, Percent, StarStar,
    LParen, RParen, Semi, Eq,
    Var, Typeof,
    True, False, Null, Undefined,
    Eof,
}

struct Lexer<'a> {
    src: &'a [u8],
    pos: usize,
}

impl<'a> Lexer<'a> {
    fn new(src: &'a str) -> Self { Lexer { src: src.as_bytes(), pos: 0 } }

    fn skip_ws(&mut self) {
        while self.pos < self.src.len() && (self.src[self.pos] == b' ' || self.src[self.pos] == b'\t' || self.src[self.pos] == b'\n' || self.src[self.pos] == b'\r') {
            self.pos += 1;
        }
    }

    fn next(&mut self) -> Tok {
        self.skip_ws();
        if self.pos >= self.src.len() { return Tok::Eof; }
        let c = self.src[self.pos];
        match c {
            b'+' => { self.pos += 1; Tok::Plus }
            b'-' => { self.pos += 1; Tok::Minus }
            b'*' => {
                self.pos += 1;
                if self.pos < self.src.len() && self.src[self.pos] == b'*' {
                    self.pos += 1; Tok::StarStar
                } else { Tok::Star }
            }
            b'/' => { self.pos += 1; Tok::Slash }
            b'%' => { self.pos += 1; Tok::Percent }
            b'(' => { self.pos += 1; Tok::LParen }
            b')' => { self.pos += 1; Tok::RParen }
            b';' => { self.pos += 1; Tok::Semi }
            b'=' => { self.pos += 1; Tok::Eq }
            b'\'' | b'"' => {
                let q = c;
                self.pos += 1;
                let start = self.pos;
                while self.pos < self.src.len() && self.src[self.pos] != q {
                    self.pos += 1;
                }
                let s = String::from_utf8_lossy(&self.src[start..self.pos]).to_string();
                if self.pos < self.src.len() { self.pos += 1; }
                Tok::Str(s)
            }
            b'0'..=b'9' | b'.' => {
                let start = self.pos;
                while self.pos < self.src.len() && (self.src[self.pos].is_ascii_digit() || self.src[self.pos] == b'.') {
                    self.pos += 1;
                }
                let s = String::from_utf8_lossy(&self.src[start..self.pos]);
                Tok::Num(s.parse().unwrap_or(0.0))
            }
            b'a'..=b'z' | b'A'..=b'Z' | b'_' | b'$' => {
                let start = self.pos;
                while self.pos < self.src.len() && (self.src[self.pos].is_ascii_alphanumeric() || self.src[self.pos] == b'_' || self.src[self.pos] == b'$') {
                    self.pos += 1;
                }
                let word = String::from_utf8_lossy(&self.src[start..self.pos]).to_string();
                match word.as_str() {
                    "var" | "let" | "const" => Tok::Var,
                    "typeof" => Tok::Typeof,
                    "true" => Tok::True,
                    "false" => Tok::False,
                    "null" => Tok::Null,
                    "undefined" => Tok::Undefined,
                    _ => Tok::Ident(word),
                }
            }
            _ => { self.pos += 1; Tok::Eof }
        }
    }

    fn peek(&mut self) -> Tok {
        let saved = self.pos;
        let t = self.next();
        self.pos = saved;
        t
    }
}

// ── Evaluator ─────────────────────────────────────────────────────────────────

struct Eval<'a> {
    lex: Lexer<'a>,
    vars: HashMap<String, Value>,
}

impl<'a> Eval<'a> {
    fn new(src: &'a str, sandbox: Option<&Value>) -> Self {
        let mut vars = HashMap::new();
        if let Some(Value::Object(obj)) = sandbox {
            let obj = obj.lock().unwrap();
            for (k, v) in &obj.properties {
                if !k.starts_with("__") {
                    vars.insert(k.clone(), v.clone());
                }
            }
        }
        Eval { lex: Lexer::new(src), vars }
    }

    fn run(&mut self) -> Value {
        let mut last = Value::Undefined;
        loop {
            if self.lex.peek() == Tok::Eof { break; }
            // Skip bare semicolons
            if self.lex.peek() == Tok::Semi { self.lex.next(); continue; }
            // var / let / const declaration
            if self.lex.peek() == Tok::Var {
                self.lex.next();
                if let Tok::Ident(name) = self.lex.next() {
                    if self.lex.peek() == Tok::Eq {
                        self.lex.next();
                        let val = self.expr();
                        self.vars.insert(name, val.clone());
                        last = val;
                    } else {
                        self.vars.insert(name, Value::Undefined);
                    }
                }
                if self.lex.peek() == Tok::Semi { self.lex.next(); }
                continue;
            }
            last = self.expr();
            if self.lex.peek() == Tok::Semi { self.lex.next(); }
        }
        last
    }

    fn expr(&mut self) -> Value { self.add() }

    fn add(&mut self) -> Value {
        let mut left = self.mul();
        loop {
            match self.lex.peek() {
                Tok::Plus => { self.lex.next(); left = val_add(left, self.mul()); }
                Tok::Minus => { self.lex.next(); left = val_sub(left, self.mul()); }
                _ => break,
            }
        }
        left
    }

    fn mul(&mut self) -> Value {
        let mut left = self.pow();
        loop {
            match self.lex.peek() {
                Tok::Star => { self.lex.next(); left = val_mul(left, self.pow()); }
                Tok::Slash => { self.lex.next(); left = val_div(left, self.pow()); }
                Tok::Percent => { self.lex.next(); left = val_rem(left, self.pow()); }
                _ => break,
            }
        }
        left
    }

    fn pow(&mut self) -> Value {
        let base = self.unary();
        if self.lex.peek() == Tok::StarStar {
            self.lex.next();
            let exp = self.pow(); // right-associative
            val_pow(base, exp)
        } else {
            base
        }
    }

    fn unary(&mut self) -> Value {
        match self.lex.peek() {
            Tok::Minus => { self.lex.next(); val_neg(self.unary()) }
            Tok::Typeof => {
                self.lex.next();
                let v = self.unary();
                s(typeof_str(&v))
            }
            _ => self.primary(),
        }
    }

    fn primary(&mut self) -> Value {
        match self.lex.next() {
            Tok::Num(n) => {
                if n.fract() == 0.0 && n >= i32::MIN as f64 && n <= i32::MAX as f64 {
                    Value::I32(n as i32)
                } else {
                    Value::F64(n)
                }
            }
            Tok::Str(s) => Value::String(Arc::from(s.as_str())),
            Tok::True => Value::Bool(true),
            Tok::False => Value::Bool(false),
            Tok::Null => Value::Null,
            Tok::Undefined => Value::Undefined,
            Tok::Ident(name) => {
                self.vars.get(&name).cloned().unwrap_or(Value::Undefined)
            }
            Tok::LParen => {
                let v = self.expr();
                if self.lex.peek() == Tok::RParen { self.lex.next(); }
                v
            }
            _ => Value::Undefined,
        }
    }
}

fn typeof_str(v: &Value) -> &'static str {
    match v {
        Value::Bool(_) => "boolean",
        Value::I32(_) | Value::F64(_) | Value::I64(_) => "number",
        Value::String(_) => "string",
        Value::Null => "object",
        Value::Undefined => "undefined",
        Value::Object(_) => "object",
        _ => "undefined",
    }
}

fn to_f64(v: &Value) -> f64 {
    match v {
        Value::I32(n) => *n as f64,
        Value::F64(f) => *f,
        Value::I64(n) => *n as f64,
        _ => f64::NAN,
    }
}

fn num_result(f: f64) -> Value {
    if f.fract() == 0.0 && f >= i32::MIN as f64 && f <= i32::MAX as f64 {
        Value::I32(f as i32)
    } else {
        Value::F64(f)
    }
}

fn val_add(a: Value, b: Value) -> Value {
    match (&a, &b) {
        (Value::String(s1), Value::String(s2)) => Value::String(Arc::from(format!("{s1}{s2}").as_str())),
        _ => num_result(to_f64(&a) + to_f64(&b)),
    }
}
fn val_sub(a: Value, b: Value) -> Value { num_result(to_f64(&a) - to_f64(&b)) }
fn val_mul(a: Value, b: Value) -> Value { num_result(to_f64(&a) * to_f64(&b)) }
fn val_div(a: Value, b: Value) -> Value { Value::F64(to_f64(&a) / to_f64(&b)) }
fn val_rem(a: Value, b: Value) -> Value { num_result(to_f64(&a) % to_f64(&b)) }
fn val_pow(a: Value, b: Value) -> Value { num_result(to_f64(&a).powf(to_f64(&b))) }
fn val_neg(a: Value) -> Value { num_result(-to_f64(&a)) }

fn eval_code(code: &str, sandbox: Option<&Value>) -> Value {
    Eval::new(code, sandbox).run()
}

// ── Registration ─────────────────────────────────────────────────────────────

pub fn register(vm: &mut VM) {
    vm.register_host_fn("node:vm", "runInNewContext", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => return Value::Undefined,
        };
        let sandbox = args.get(1);
        eval_code(&code, sandbox)
    }));

    vm.register_host_fn("node:vm", "runInThisContext", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => return Value::Undefined,
        };
        eval_code(&code, None)
    }));

    vm.register_host_fn("node:vm", "createContext", Box::new(|_ctx, args| {
        match args.first() {
            Some(Value::Object(obj)) => {
                obj.lock().unwrap().properties.insert("__isContext".into(), Value::Bool(true));
                Value::Object(Arc::clone(obj))
            }
            _ => {
                let mut o = Object::new();
                o.properties.insert("__isContext".into(), Value::Bool(true));
                Value::Object(Arc::new(Mutex::new(o)))
            }
        }
    }));

    vm.register_host_fn("node:vm", "isContext", Box::new(|_ctx, args| {
        match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                Value::Bool(matches!(o.properties.get("__isContext"), Some(Value::Bool(true))))
            }
            _ => Value::Bool(false),
        }
    }));

    vm.register_host_fn("node:vm", "Script", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => String::new(),
        };
        let mut o = Object::new();
        o.properties.insert("__code".into(), Value::String(Arc::from(code.as_str())));
        Value::Object(Arc::new(Mutex::new(o)))
    }));

    vm.register_host_fn("node:vm", "scriptRunInNewContext", Box::new(|_ctx, args| {
        let script = args.first();
        let code = match script {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                match o.properties.get("__code") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => return Value::Undefined,
                }
            }
            _ => return Value::Undefined,
        };
        let sandbox = args.get(1);
        eval_code(&code, sandbox)
    }));

    vm.register_host_fn("node:vm", "scriptRunInThisContext", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::Object(obj)) => {
                let o = obj.lock().unwrap();
                match o.properties.get("__code") {
                    Some(Value::String(s)) => s.to_string(),
                    _ => return Value::Undefined,
                }
            }
            _ => return Value::Undefined,
        };
        eval_code(&code, None)
    }));

    vm.register_host_fn("node:vm", "compileFunction", Box::new(|_ctx, args| {
        let code = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            _ => String::new(),
        };
        let mut o = Object::new();
        o.properties.insert("__code".into(), Value::String(Arc::from(code.as_str())));
        Value::Object(Arc::new(Mutex::new(o)))
    }));

    vm.register_host_fn("node:vm", "measureMemory", Box::new(|_ctx, _args| {
        let mut o = Object::new();
        o.properties.insert("total".into(), Value::F64(0.0));
        o.properties.insert("jsMemoryEstimate".into(), Value::F64(0.0));
        Value::Object(Arc::new(Mutex::new(o)))
    }));
}
