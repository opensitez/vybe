//! PHP's secretbox surface. PHP strings containing binary data are represented
//! by one Latin-1 code point per byte in the VM (as in `random_bytes`).
//! Secretbox is XSalsa20-Poly1305 with a 24-byte nonce and a 32-byte key.

use std::sync::Arc;
use vybe_runtime::{Framework, Value};

fn bytes(value: &Value) -> Option<Vec<u8>> {
    let Value::String(s) = value else { return None };
    s.chars().map(|ch| u8::try_from(ch as u32).ok()).collect()
}

fn binary_string(data: &[u8]) -> Value {
    Value::String(Arc::from(
        data.iter().copied().map(char::from).collect::<String>(),
    ))
}

fn word(data: &[u8]) -> u32 {
    u32::from_le_bytes(data.try_into().unwrap())
}

fn quarter(x: &mut [u32; 16], a: usize, b: usize, c: usize, d: usize) {
    x[b] ^= x[a].wrapping_add(x[d]).rotate_left(7);
    x[c] ^= x[b].wrapping_add(x[a]).rotate_left(9);
    x[d] ^= x[c].wrapping_add(x[b]).rotate_left(13);
    x[a] ^= x[d].wrapping_add(x[c]).rotate_left(18);
}

fn rounds(x: &mut [u32; 16]) {
    for _ in 0..10 {
        quarter(x, 0, 4, 8, 12);
        quarter(x, 5, 9, 13, 1);
        quarter(x, 10, 14, 2, 6);
        quarter(x, 15, 3, 7, 11);
        quarter(x, 0, 1, 2, 3);
        quarter(x, 5, 6, 7, 4);
        quarter(x, 10, 11, 8, 9);
        quarter(x, 15, 12, 13, 14);
    }
}

fn initial_state(key: &[u8; 32], input: &[u8; 16]) -> [u32; 16] {
    let sigma = b"expand 32-byte k";
    let mut state = [0u32; 16];
    for (slot, offset) in [(0, 0), (5, 4), (10, 8), (15, 12)] {
        state[slot] = word(&sigma[offset..offset + 4]);
    }
    for i in 0..4 {
        state[1 + i] = word(&key[i * 4..i * 4 + 4]);
        state[11 + i] = word(&key[16 + i * 4..20 + i * 4]);
        state[6 + i] = word(&input[i * 4..i * 4 + 4]);
    }
    state
}

fn hsalsa20(key: &[u8; 32], nonce: &[u8; 16]) -> [u8; 32] {
    let mut state = initial_state(key, nonce);
    rounds(&mut state);
    let mut out = [0u8; 32];
    for (i, index) in [0, 5, 10, 15, 6, 7, 8, 9].into_iter().enumerate() {
        out[i * 4..i * 4 + 4].copy_from_slice(&state[index].to_le_bytes());
    }
    out
}

fn salsa20_block(key: &[u8; 32], nonce: &[u8; 8], counter: u64) -> [u8; 64] {
    let mut input = [0u8; 16];
    input[..8].copy_from_slice(nonce);
    input[8..].copy_from_slice(&counter.to_le_bytes());
    let state = initial_state(key, &input);
    let mut output = state;
    rounds(&mut output);
    let mut bytes = [0u8; 64];
    for i in 0..16 {
        bytes[i * 4..i * 4 + 4].copy_from_slice(&output[i].wrapping_add(state[i]).to_le_bytes());
    }
    bytes
}

fn xor_message(message: &[u8], key: &[u8; 32], nonce: &[u8; 8]) -> Vec<u8> {
    let mut result = Vec::with_capacity(message.len());
    let mut position = 0usize;
    let mut counter = 0u64;
    while position < message.len() {
        let block = salsa20_block(key, nonce, counter);
        let start = if counter == 0 { 32 } else { 0 };
        let count = (64 - start).min(message.len() - position);
        for i in 0..count {
            result.push(message[position + i] ^ block[start + i]);
        }
        position += count;
        counter += 1;
    }
    result
}

fn poly1305(message: &[u8], key: &[u8; 32]) -> [u8; 16] {
    let mut r_bytes = [0u8; 16];
    r_bytes.copy_from_slice(&key[..16]);
    for index in [3, 7, 11, 15] {
        r_bytes[index] &= 15;
    }
    for index in [4, 8, 12] {
        r_bytes[index] &= 252;
    }
    let r = limbs(&r_bytes);
    let mut h = [0u64; 5];
    for part in message.chunks(16) {
        let mut block = [0u8; 16];
        block[..part.len()].copy_from_slice(part);
        if part.len() < 16 {
            block[part.len()] = 1;
        }
        let m = limbs(&block);
        for i in 0..5 {
            h[i] += m[i];
        }
        if part.len() == 16 {
            h[4] += 1 << 24;
        }
        let d = [
            h[0] * r[0] + 5 * (h[1] * r[4] + h[2] * r[3] + h[3] * r[2] + h[4] * r[1]),
            h[0] * r[1] + h[1] * r[0] + 5 * (h[2] * r[4] + h[3] * r[3] + h[4] * r[2]),
            h[0] * r[2] + h[1] * r[1] + h[2] * r[0] + 5 * (h[3] * r[4] + h[4] * r[3]),
            h[0] * r[3] + h[1] * r[2] + h[2] * r[1] + h[3] * r[0] + 5 * h[4] * r[4],
            h[0] * r[4] + h[1] * r[3] + h[2] * r[2] + h[3] * r[1] + h[4] * r[0],
        ];
        let mask = (1u64 << 26) - 1;
        let mut carry = 0u64;
        for i in 0..5 {
            let value = d[i] + carry;
            h[i] = value & mask;
            carry = value >> 26;
        }
        h[0] += carry * 5;
        carry = h[0] >> 26;
        h[0] &= mask;
        h[1] += carry;
    }

    let mask = (1u64 << 26) - 1;
    for i in 0..4 {
        let carry = h[i] >> 26;
        h[i] &= mask;
        h[i + 1] += carry;
    }
    let carry = h[4] >> 26;
    h[4] &= mask;
    h[0] += carry * 5;
    let carry = h[0] >> 26;
    h[0] &= mask;
    h[1] += carry;

    let mut g = h;
    g[0] += 5;
    for i in 0..4 {
        let carry = g[i] >> 26;
        g[i] &= mask;
        g[i + 1] += carry;
    }
    if g[4] >= 1 << 26 {
        g[4] -= 1 << 26;
        h = g;
    }

    let words = [
        h[0] | (h[1] << 26),
        (h[1] >> 6) | (h[2] << 20),
        (h[2] >> 12) | (h[3] << 14),
        (h[3] >> 18) | (h[4] << 8),
    ];
    let mut tag = [0u8; 16];
    let mut carry = 0u64;
    for i in 0..4 {
        let value = (words[i] & 0xffff_ffff) + word(&key[16 + i * 4..20 + i * 4]) as u64 + carry;
        tag[i * 4..i * 4 + 4].copy_from_slice(&(value as u32).to_le_bytes());
        carry = value >> 32;
    }
    tag
}

fn limbs(data: &[u8; 16]) -> [u64; 5] {
    let get = |offset| word(&data[offset..offset + 4]) as u64;
    let mask = (1u64 << 26) - 1;
    [
        get(0) & mask,
        (get(3) >> 2) & mask,
        (get(6) >> 4) & mask,
        (get(9) >> 6) & mask,
        (get(12) >> 8) & mask,
    ]
}

fn secretbox(message: &[u8], nonce: &[u8; 24], key: &[u8; 32]) -> Vec<u8> {
    let subkey = hsalsa20(key, nonce[..16].try_into().unwrap());
    let tail: &[u8; 8] = nonce[16..].try_into().unwrap();
    let first = salsa20_block(&subkey, tail, 0);
    let ciphertext = xor_message(message, &subkey, tail);
    let tag = poly1305(&ciphertext, first[..32].try_into().unwrap());
    let mut result = Vec::with_capacity(16 + ciphertext.len());
    result.extend_from_slice(&tag);
    result.extend_from_slice(&ciphertext);
    result
}

fn secretbox_open(ciphertext: &[u8], nonce: &[u8; 24], key: &[u8; 32]) -> Option<Vec<u8>> {
    if ciphertext.len() < 16 {
        return None;
    }
    let subkey = hsalsa20(key, nonce[..16].try_into().unwrap());
    let tail: &[u8; 8] = nonce[16..].try_into().unwrap();
    let first = salsa20_block(&subkey, tail, 0);
    let expected = poly1305(&ciphertext[16..], first[..32].try_into().unwrap());
    let difference = expected
        .iter()
        .zip(&ciphertext[..16])
        .fold(0u8, |acc, (a, b)| acc | (a ^ b));
    if difference != 0 {
        return None;
    }
    Some(xor_message(&ciphertext[16..], &subkey, tail))
}

pub fn register(fw: &mut Framework<'_>) {
    fw.register_host_fn(
        "php:sodium",
        "keygen",
        Box::new(|ctx, _| {
            let mut key = [0u8; 32];
            if getrandom::getrandom(&mut key).is_err() {
                ctx.throw_value(Value::String(Arc::from(
                    "sodium_crypto_secretbox_keygen(): random source failed",
                )));
                return Value::Null;
            }
            binary_string(&key)
        }),
    );
    fw.register_host_fn(
        "php:sodium",
        "secretbox",
        Box::new(|ctx, args| {
            let (Some(message), Some(nonce), Some(key)) = (
                args.first().and_then(bytes),
                args.get(1).and_then(bytes),
                args.get(2).and_then(bytes),
            ) else {
                return Value::Bool(false);
            };
            if nonce.len() != 24 || key.len() != 32 {
                ctx.throw_value(Value::String(Arc::from(
                    "sodium_crypto_secretbox(): invalid nonce or key length",
                )));
                return Value::Null;
            }
            binary_string(&secretbox(
                &message,
                nonce.as_slice().try_into().unwrap(),
                key.as_slice().try_into().unwrap(),
            ))
        }),
    );
    fw.register_host_fn(
        "php:sodium",
        "secretbox_open",
        Box::new(|ctx, args| {
            let (Some(ciphertext), Some(nonce), Some(key)) = (
                args.first().and_then(bytes),
                args.get(1).and_then(bytes),
                args.get(2).and_then(bytes),
            ) else {
                return Value::Bool(false);
            };
            if nonce.len() != 24 || key.len() != 32 {
                ctx.throw_value(Value::String(Arc::from(
                    "sodium_crypto_secretbox_open(): invalid nonce or key length",
                )));
                return Value::Null;
            }
            match secretbox_open(
                &ciphertext,
                nonce.as_slice().try_into().unwrap(),
                key.as_slice().try_into().unwrap(),
            ) {
                Some(plaintext) => binary_string(&plaintext),
                None => Value::Bool(false),
            }
        }),
    );
}
