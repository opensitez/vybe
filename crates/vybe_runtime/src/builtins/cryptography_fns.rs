use crate::value::{Value, RuntimeError};
use md5;
use sha2::{Sha256, Digest};

/// Compute MD5 hash of a byte array or string
pub fn md5_hash_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("MD5.ComputeHash requires data".to_string()));
    }
    
    let bytes = value_to_bytes(&args[0]);
    let digest = md5::compute(&bytes);
    
    // Return as Byte array (standard .NET behavior)
    let hash_bytes = digest.0.iter().map(|b| Value::Byte(*b)).collect();
    Ok(Value::Array(hash_bytes))
}

/// Compute SHA256 hash of a byte array or string
pub fn sha256_hash_fn(args: &[Value]) -> Result<Value, RuntimeError> {
    if args.is_empty() {
        return Err(RuntimeError::Custom("SHA256.ComputeHash requires data".to_string()));
    }
    
    let bytes = value_to_bytes(&args[0]);
    let mut hasher = Sha256::new();
    hasher.update(&bytes);
    let result = hasher.finalize();
    
    let hash_bytes = result.iter().map(|b| Value::Byte(*b)).collect();
    Ok(Value::Array(hash_bytes))
}

fn value_to_bytes(val: &Value) -> Vec<u8> {
    match val {
        Value::Array(arr) => arr.iter().map(|v| match v {
            Value::Byte(b) => *b,
            Value::Integer(i) => *i as u8,
            _ => 0u8,
        }).collect(),
        Value::String(s) => s.as_bytes().to_vec(),
        _ => val.as_string().as_bytes().to_vec(),
    }
}
