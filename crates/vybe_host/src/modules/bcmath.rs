//! PHP BCMath host surface.
//!
//! BCMath is arbitrary precision decimal-string arithmetic. It is not covered
//! by libc math, WASM numeric opcodes, or Vybe's ECMA BigInt shim (currently
//! bounded to i64), so this module keeps the PHP extension surface as a pure
//! computation import.

use std::cmp::Ordering;
use std::sync::{Arc, Mutex};
use vybe_bytecode::{HostContext, VM, Value};

#[derive(Clone, Debug)]
struct Dec {
    neg: bool,
    digits: Vec<u8>,
    scale: usize,
}

impl Dec {
    fn zero() -> Self {
        Self {
            neg: false,
            digits: vec![0],
            scale: 0,
        }
    }

    fn parse(text: &str) -> Self {
        let mut s = text.trim();
        let mut neg = false;
        if let Some(rest) = s.strip_prefix('-') {
            neg = true;
            s = rest;
        } else if let Some(rest) = s.strip_prefix('+') {
            s = rest;
        }

        let mut digits = Vec::new();
        let mut scale = 0usize;
        let mut seen_dot = false;
        for ch in s.chars() {
            if ch == '.' && !seen_dot {
                seen_dot = true;
                continue;
            }
            if let Some(d) = ch.to_digit(10) {
                digits.push(d as u8);
                if seen_dot {
                    scale += 1;
                }
            } else {
                break;
            }
        }
        if digits.is_empty() {
            return Self::zero();
        }
        let mut out = Self { neg, digits, scale };
        out.normalize();
        out
    }

    fn normalize(&mut self) {
        while self.digits.len() > self.scale + 1 && self.digits.first() == Some(&0) {
            self.digits.remove(0);
        }
        if self.digits.iter().all(|d| *d == 0) {
            self.neg = false;
            self.digits = vec![0];
            self.scale = 0;
        }
    }

    fn to_scale(&self, scale: usize) -> Self {
        let mut out = self.clone();
        if out.scale < scale {
            out.digits
                .extend(std::iter::repeat(0).take(scale - out.scale));
            out.scale = scale;
        } else if out.scale > scale {
            let drop = out.scale - scale;
            if drop >= out.digits.len() {
                out.digits = vec![0];
            } else {
                out.digits.truncate(out.digits.len() - drop);
            }
            out.scale = scale;
        }
        out.normalize();
        out
    }

    fn render(&self, scale: usize) -> String {
        let n = self.to_scale(scale);
        let mut text = String::new();
        if n.neg && !is_zero_digits(&n.digits) {
            text.push('-');
        }
        if scale == 0 {
            for d in n.digits {
                text.push(char::from(b'0' + d));
            }
            return text;
        }
        if n.digits.len() <= scale {
            text.push('0');
            text.push('.');
            for _ in 0..(scale - n.digits.len()) {
                text.push('0');
            }
            for d in n.digits {
                text.push(char::from(b'0' + d));
            }
        } else {
            let split = n.digits.len() - scale;
            for d in &n.digits[..split] {
                text.push(char::from(b'0' + *d));
            }
            text.push('.');
            for d in &n.digits[split..] {
                text.push(char::from(b'0' + *d));
            }
        }
        text
    }
}

fn is_zero_digits(digits: &[u8]) -> bool {
    digits.iter().all(|d| *d == 0)
}

fn val_string(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{v}")).unwrap_or_default()
}

fn val_scale(args: &[Value], idx: usize, default_scale: &Arc<Mutex<i32>>) -> usize {
    args.get(idx)
        .map(|v| v.as_f64() as i32)
        .unwrap_or_else(|| *default_scale.lock().unwrap())
        .max(0) as usize
}

fn s_val(text: String) -> Value {
    Value::String(Arc::from(text.as_str()))
}

fn cmp_abs_digits(a: &[u8], b: &[u8]) -> Ordering {
    let a0 = a.iter().position(|d| *d != 0).unwrap_or(a.len());
    let b0 = b.iter().position(|d| *d != 0).unwrap_or(b.len());
    let aa = &a[a0..];
    let bb = &b[b0..];
    aa.len().cmp(&bb.len()).then_with(|| aa.cmp(bb))
}

fn add_digits(a: &[u8], b: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    let mut carry = 0u8;
    let mut ia = a.len();
    let mut ib = b.len();
    while ia > 0 || ib > 0 || carry > 0 {
        let da = if ia > 0 {
            ia -= 1;
            a[ia]
        } else {
            0
        };
        let db = if ib > 0 {
            ib -= 1;
            b[ib]
        } else {
            0
        };
        let sum = da + db + carry;
        out.push(sum % 10);
        carry = sum / 10;
    }
    out.reverse();
    out
}

fn sub_digits(a: &[u8], b: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    let mut borrow = 0i8;
    let mut ia = a.len();
    let mut ib = b.len();
    while ia > 0 {
        ia -= 1;
        let mut da = a[ia] as i8 - borrow;
        let db = if ib > 0 {
            ib -= 1;
            b[ib] as i8
        } else {
            0
        };
        if da < db {
            da += 10;
            borrow = 1;
        } else {
            borrow = 0;
        }
        out.push((da - db) as u8);
    }
    out.reverse();
    trim_digits(out)
}

fn trim_digits(mut digits: Vec<u8>) -> Vec<u8> {
    while digits.len() > 1 && digits.first() == Some(&0) {
        digits.remove(0);
    }
    digits
}

fn align(mut a: Dec, mut b: Dec) -> (Dec, Dec) {
    let scale = a.scale.max(b.scale);
    if a.scale < scale {
        a.digits.extend(std::iter::repeat(0).take(scale - a.scale));
        a.scale = scale;
    }
    if b.scale < scale {
        b.digits.extend(std::iter::repeat(0).take(scale - b.scale));
        b.scale = scale;
    }
    (a, b)
}

fn add_dec(a: Dec, b: Dec) -> Dec {
    let (a, b) = align(a, b);
    if a.neg == b.neg {
        let mut out = Dec {
            neg: a.neg,
            digits: add_digits(&a.digits, &b.digits),
            scale: a.scale,
        };
        out.normalize();
        return out;
    }
    match cmp_abs_digits(&a.digits, &b.digits) {
        Ordering::Greater | Ordering::Equal => {
            let mut out = Dec {
                neg: a.neg,
                digits: sub_digits(&a.digits, &b.digits),
                scale: a.scale,
            };
            out.normalize();
            out
        }
        Ordering::Less => {
            let mut out = Dec {
                neg: b.neg,
                digits: sub_digits(&b.digits, &a.digits),
                scale: a.scale,
            };
            out.normalize();
            out
        }
    }
}

fn mul_dec(a: Dec, b: Dec) -> Dec {
    if is_zero_digits(&a.digits) || is_zero_digits(&b.digits) {
        return Dec::zero();
    }
    let mut out = vec![0u16; a.digits.len() + b.digits.len()];
    for (ia, da) in a.digits.iter().rev().enumerate() {
        for (ib, db) in b.digits.iter().rev().enumerate() {
            let idx = out.len() - 1 - ia - ib;
            out[idx] += *da as u16 * *db as u16;
        }
    }
    for i in (1..out.len()).rev() {
        let carry = out[i] / 10;
        out[i] %= 10;
        out[i - 1] += carry;
    }
    let mut dec = Dec {
        neg: a.neg ^ b.neg,
        digits: trim_digits(out.into_iter().map(|d| d as u8).collect()),
        scale: a.scale + b.scale,
    };
    dec.normalize();
    dec
}

fn div_digits(numer: Vec<u8>, denom: Vec<u8>) -> Vec<u8> {
    if is_zero_digits(&denom) {
        return vec![0];
    }
    let mut q = Vec::new();
    let mut rem: Vec<u8> = Vec::new();
    for d in trim_digits(numer) {
        rem.push(d);
        rem = trim_digits(rem);
        let mut digit = 0u8;
        while cmp_abs_digits(&rem, &denom) != Ordering::Less {
            rem = sub_digits(&rem, &denom);
            digit += 1;
        }
        q.push(digit);
    }
    trim_digits(q)
}

fn div_dec(a: Dec, b: Dec, scale: usize) -> Dec {
    if is_zero_digits(&b.digits) {
        return Dec::zero();
    }
    let mut numer = a.digits.clone();
    numer.extend(std::iter::repeat(0).take(b.scale + scale));
    let mut denom = b.digits.clone();
    denom.extend(std::iter::repeat(0).take(a.scale));
    let mut out = Dec {
        neg: a.neg ^ b.neg,
        digits: div_digits(numer, denom),
        scale,
    };
    out.normalize();
    out
}

fn cmp_dec(a: Dec, b: Dec, scale: usize) -> Ordering {
    let (a, b) = align(a.to_scale(scale), b.to_scale(scale));
    if a.neg != b.neg {
        return if a.neg {
            Ordering::Less
        } else {
            Ordering::Greater
        };
    }
    let ord = cmp_abs_digits(&a.digits, &b.digits);
    if a.neg { ord.reverse() } else { ord }
}

fn pow_dec(base: Dec, exp: i64, scale: usize) -> Dec {
    if exp < 0 {
        return Dec::zero();
    }
    let mut result = Dec::parse("1");
    let mut b = base;
    let mut e = exp;
    while e > 0 {
        if e & 1 == 1 {
            result = mul_dec(result, b.clone()).to_scale(scale);
        }
        e >>= 1;
        if e > 0 {
            b = mul_dec(b.clone(), b).to_scale(scale);
        }
    }
    result
}

fn sqrt_dec(value: Dec, scale: usize) -> Dec {
    let n = value.render(scale + 8).parse::<f64>().unwrap_or(0.0).sqrt();
    Dec::parse(&format!("{:.*}", scale + 4, n)).to_scale(scale)
}

pub fn register(vm: &mut VM) {
    let default_scale = Arc::new(Mutex::new(0i32));

    macro_rules! binop {
        ($name:literal, $op:expr) => {{
            let scale_ref = default_scale.clone();
            vm.register_host_fn(
                "php:bcmath",
                $name,
                Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
                    let scale = val_scale(args, 2, &scale_ref);
                    let a = Dec::parse(&val_string(args, 0));
                    let b = Dec::parse(&val_string(args, 1));
                    let result: Dec = $op(a, b, scale);
                    s_val(result.render(scale))
                }),
            );
        }};
    }

    binop!("bcadd", |a: Dec, b: Dec, _scale| add_dec(a, b));
    binop!("bcsub", |a: Dec, mut b: Dec, _scale| {
        b.neg = !b.neg;
        add_dec(a, b)
    });
    binop!("bcmul", |a: Dec, b: Dec, scale| mul_dec(a, b)
        .to_scale(scale));
    binop!("bcdiv", |a: Dec, b: Dec, scale| div_dec(a, b, scale));
    binop!("bcmod", |a: Dec, b: Dec, scale| {
        let quotient = div_dec(a.clone(), b.clone(), 0);
        let product = mul_dec(quotient, b);
        let mut neg_product = product;
        neg_product.neg = !neg_product.neg;
        add_dec(a, neg_product).to_scale(scale)
    });

    let scale_ref = default_scale.clone();
    vm.register_host_fn(
        "php:bcmath",
        "bcpow",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let scale = val_scale(args, 2, &scale_ref);
            let base = Dec::parse(&val_string(args, 0));
            let exp = val_string(args, 1).parse::<i64>().unwrap_or(0);
            s_val(pow_dec(base, exp, scale).render(scale))
        }),
    );

    let scale_ref = default_scale.clone();
    vm.register_host_fn(
        "php:bcmath",
        "bcsqrt",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let scale = val_scale(args, 1, &scale_ref);
            let value = Dec::parse(&val_string(args, 0));
            s_val(sqrt_dec(value, scale).render(scale))
        }),
    );

    let scale_ref = default_scale.clone();
    vm.register_host_fn(
        "php:bcmath",
        "bccomp",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let a = Dec::parse(&val_string(args, 0));
            let b = Dec::parse(&val_string(args, 1));
            let scale = args
                .get(2)
                .map(|v| (v.as_f64() as i32).max(0) as usize)
                .unwrap_or_else(|| {
                    let global = *scale_ref.lock().unwrap();
                    if global > 0 {
                        global as usize
                    } else {
                        a.scale.max(b.scale)
                    }
                });
            let result = match cmp_dec(a, b, scale) {
                Ordering::Less => -1.0,
                Ordering::Equal => 0.0,
                Ordering::Greater => 1.0,
            };
            Value::F64(result)
        }),
    );

    vm.register_host_fn(
        "php:bcmath",
        "bcscale",
        Box::new(move |_ctx: &mut HostContext, args: &[Value]| {
            let old = *default_scale.lock().unwrap();
            if let Some(value) = args.first() {
                *default_scale.lock().unwrap() = (value.as_f64() as i32).max(0);
            }
            Value::F64(old as f64)
        }),
    );
}
