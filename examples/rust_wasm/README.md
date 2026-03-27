# Rust → WASM → Vybe Example

Compile Rust to WASM, then call it from VB or JS through the Vybe VM.

## Setup

```bash
# Install rustup if not already (needed for wasm target)
brew install rustup-init && rustup-init

# Add WASM target
rustup target add wasm32-unknown-unknown

# Compile to WASM
cargo build --release --target wasm32-unknown-unknown

# The .wasm file is at:
# target/wasm32-unknown-unknown/release/vybe_rust_wasm.wasm

# Copy to examples
cp target/wasm32-unknown-unknown/release/vybe_rust_wasm.wasm ./math.wasm
```

## Run from Vybe

```bash
# Load the WASM module (when native WASM loading is implemented)
vybec load_rust.vybe
```

## What it demonstrates

- Pure Rust functions (add, factorial, fibonacci, is_prime) compiled to WASM
- No runtime dependencies — just math
- Can be called from VB, JS, or any WASM runtime
