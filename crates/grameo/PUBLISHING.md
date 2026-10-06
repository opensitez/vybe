# Publishing Grameo

Packages: `grameo` and `grameo_codegen`, version `0.1.0`. Runtime source: `crates/grameo`; generator source: `crates/grameo/codegen`.

From the repository root:

```sh
cargo test --manifest-path crates/grameo/Cargo.toml --workspace --offline
cargo clippy --manifest-path crates/grameo/Cargo.toml --workspace --all-targets --offline -- -D warnings
cargo package --manifest-path crates/grameo/Cargo.toml -p grameo --allow-dirty --offline
```

Before the first release, verify the generator against the unpacked runtime archive:

```sh
cargo package --manifest-path crates/grameo/codegen/Cargo.toml --allow-dirty --offline --config 'patch.crates-io.grameo.path="crates/grameo/target/package/grameo-0.1.0"'
```

Publish the runtime first. Once it is available on crates.io, the generator can resolve its versioned dependency normally:

```sh
cargo publish --manifest-path crates/grameo/Cargo.toml -p grameo --dry-run
cargo publish --manifest-path crates/grameo/Cargo.toml -p grameo
cargo publish --manifest-path crates/grameo/codegen/Cargo.toml --dry-run
cargo publish --manifest-path crates/grameo/codegen/Cargo.toml
```

Publishing requires registry credentials. Package archives include source, standalone tests, examples and license terms; repository-only corpus checks, experiment reports and build output are excluded. Inspect contents with `cargo package --list` before publishing.
