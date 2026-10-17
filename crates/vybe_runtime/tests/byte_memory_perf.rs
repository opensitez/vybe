//! Actual public scalar-byte helpers, including their synchronization overhead.
use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::shared_memory::SharedMemory;

#[test]
fn byte_helpers_preserve_alias_visibility_and_bounds() {
    let memory = SharedMemory::from_vec(vec![0x31; 32]);
    let alias = memory.clone();
    for address in 0..32 {
        memory.store_u8(address, address as u8).unwrap();
        assert_eq!(alias.load_u8(address).unwrap(), address as u8);
    }
    let before = memory.with_buffer(|bytes| bytes.to_vec());
    for address in [32, 33, usize::MAX] {
        assert!(memory.load_u8(address).is_err());
        assert!(memory.store_u8(address, 0xff).is_err());
        assert_eq!(alias.with_buffer(|bytes| bytes.to_vec()), before);
    }
    let empty = SharedMemory::from_vec(Vec::new());
    assert!(empty.load_u8(0).is_err());
    assert!(empty.store_u8(0, 1).is_err());
}

#[test]
#[ignore = "scalar-byte helper benchmark; excludes VM dispatch and module setup"]
fn loaded_scalar_byte_access_benchmark() {
    for shared in [false, true] {
        let memory = SharedMemory::from_vec(vec![0; 4096]);
        let alias = shared.then(|| memory.clone());
        let mut samples = Vec::new();
        for round in 0..23 {
            let mut checksum = 0u64;
            let start = Instant::now();
            for iteration in 0..1_000_000usize {
                let address = black_box(iteration & 4095);
                let value = black_box(iteration as u8);
                let memory = black_box(&memory);
                memory.store_u8(address, value).unwrap();
                checksum += memory.load_u8(address).unwrap() as u64;
            }
            let elapsed = start.elapsed().as_secs_f64();
            assert_eq!(checksum, (0..1_000_000usize).map(|i| i as u8 as u64).sum());
            if let Some(alias) = &alias {
                assert_eq!(alias.load_u8(0).unwrap(), memory.load_u8(0).unwrap());
            }
            if round >= 2 { samples.push(elapsed); }
        }
        samples.sort_by(f64::total_cmp);
        eprintln!("scalar-byte shared={shared} p50_ms={:.3} p95_ms={:.3}",
            samples[10] * 1000.0, samples[19] * 1000.0);
    }
}
