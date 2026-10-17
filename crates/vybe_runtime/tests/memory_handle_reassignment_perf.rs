use std::hint::black_box;
use std::time::Instant;
use vybe_runtime::shared_memory::SharedMemory;

fn sample(
    sources: &[SharedMemory; 2],
    destination: &mut SharedMemory,
    reuse: bool,
    alternating: bool,
) -> f64 {
    let start = Instant::now();
    for i in 0usize..1_000_000 {
        let index = if alternating { i & 1 } else { 0 };
        let source = black_box(&sources[index]);
        let destination = black_box(&mut *destination);
        if reuse {
            destination.clone_from(source);
        } else {
            // Clone's former default clone_from implementation.
            *destination = source.clone();
        }
    }
    start.elapsed().as_secs_f64()
}

#[test]
#[ignore = "paired loaded-machine memory-handle assignment benchmark"]
fn same_and_different_backing_assignment() {
    for alternating in [false, true] {
        let sources = [SharedMemory::new(65536), SharedMemory::new(65536)];
        sources[0].store_u8(17, 0x91).unwrap();
        sources[1].store_u8(17, 0xa5).unwrap();
        let mut previous = sources[0].clone();
        let mut current = sources[0].clone();
        let mut ratios = [0.0f64; 21];
        for pair in 0usize..23 {
            let (old, new) = if pair % 2 == 0 {
                (
                    sample(&sources, &mut previous, false, alternating),
                    sample(&sources, &mut current, true, alternating),
                )
            } else {
                let new = sample(&sources, &mut current, true, alternating);
                (sample(&sources, &mut previous, false, alternating), new)
            };
            if pair >= 2 {
                ratios[pair - 2] = old / new;
            }
            assert_eq!(current.len(), previous.len());
            let expected = if alternating { 0xa5 } else { 0x91 };
            assert_eq!(current.load_u8(17).unwrap(), expected);
            assert_eq!(previous.load_u8(17).unwrap(), expected);
        }
        ratios.sort_by(f64::total_cmp);
        eprintln!(
            "memory assignment alternating={alternating}: previous/current median {:.3}x; pairs {ratios:?}",
            ratios[10],
        );
    }
}
