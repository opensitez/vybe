use super::helpers::{compile_ok, run_prints};

// ── Coarray declarations ───────────────────────────────────────

#[test] fn coarray_scalar_decl() {
    let out = run_prints(r#"
program test
    integer :: x[*]
    x = 42
    print *, x
end program test
"#);
    assert_eq!(out, vec!["42"]);
}

#[test] fn coarray_real_decl() {
    compile_ok(r#"
program test
    real :: r[*]
    r = 3.14
    print *, r
end program test
"#);
}

#[test] fn coarray_array_decl() {
    compile_ok(r#"
program test
    integer :: a(10)[*]
    a = 0
    a(1) = 1
    print *, a(1)
end program test
"#);
}

#[test] fn coarray_2d_array_decl() {
    compile_ok(r#"
program test
    real :: m(4,4)[*]
    m = 0.0
    m(1,1) = 1.0
    print *, m(1,1)
end program test
"#);
}

#[test] fn coarray_logical_decl() {
    compile_ok(r#"
program test
    logical :: flag[*]
    flag = .true.
    print *, flag
end program test
"#);
}

#[test] fn coarray_character_decl() {
    compile_ok(r#"
program test
    character(len=20) :: msg[*]
    msg = 'hello'
    print *, trim(msg)
end program test
"#);
}

// ── THIS_IMAGE() intrinsic ────────────────────────────────────

#[test] fn this_image_basic() {
    let out = run_prints(r#"
program test
    integer :: me
    me = this_image()
    print *, me
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

#[test] fn this_image_in_conditional() {
    let out = run_prints(r#"
program test
    if (this_image() == 1) then
        print *, 'image 1'
    end if
end program test
"#);
    assert_eq!(out, vec!["image 1"]);
}

#[test] fn this_image_with_coarray() {
    compile_ok(r#"
program test
    integer :: x[*]
    integer :: coindices(1)
    x = this_image()
    coindices = this_image(x, 1)
    print *, coindices(1)
end program test
"#);
}

// ── NUM_IMAGES() intrinsic ────────────────────────────────────

#[test] fn num_images_basic() {
    let out = run_prints(r#"
program test
    print *, num_images()
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

#[test] fn num_images_requested() {
    let out = run_prints(r#"
program test
    print *, num_images(requested=.true.)
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

// ── IMAGE_INDEX() intrinsic ───────────────────────────────────

#[test] fn image_index_basic() {
    compile_ok(r#"
program test
    integer :: x[2,*]
    integer :: sub(2) = [1, 1]
    print *, image_index(x, sub)
end program test
"#);
}

#[test] fn image_index_multi_dim() {
    compile_ok(r#"
program test
    integer :: x[3,4,*]
    integer :: sub(3) = [2, 3, 1]
    print *, image_index(x, sub)
end program test
"#);
}

// ── SYNC ALL ─────────────────────────────────────────────────

#[test] fn sync_all_basic() {
    let out = run_prints(r#"
program test
    integer :: x[*]
    x = this_image() * 10
    sync all
    print *, x
end program test
"#);
    assert_eq!(out, vec!["10"]);
}

#[test] fn sync_all_with_stat() {
    let out = run_prints(r#"
program test
    integer :: stat
    sync all (stat=stat)
    if (stat /= 0) print *, 'sync error'
    print *, 'synced'
end program test
"#);
    assert_eq!(out, vec!["synced"]);
}

// ── SYNC IMAGES ──────────────────────────────────────────────

#[test] fn sync_images_star() {
    let out = run_prints(r#"
program test
    sync images (*)
    print *, 'done'
end program test
"#);
    assert_eq!(out, vec!["done"]);
}

#[test] fn sync_images_specific() {
    compile_ok(r#"
program test
    if (this_image() == 1) then
        sync images ([2, 3])
    else
        sync images ([1])
    end if
    print *, 'synced'
end program test
"#);
}

// ── SYNC MEMORY ──────────────────────────────────────────────

#[test] fn sync_memory_basic() {
    let out = run_prints(r#"
program test
    integer :: x[*]
    x = 0
    sync memory
    x = 1
    print *, x
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

// ── SYNC TEAM ────────────────────────────────────────────────

#[test] fn sync_team_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    t = get_team()
    sync team (t)
    print *, 'ok'
end program test
"#);
}

// ── CO_SUM collective ─────────────────────────────────────────

#[test] fn co_sum_scalar_int() {
    let out = run_prints(r#"
program test
    integer :: x
    x = this_image()
    call co_sum(x)
    if (this_image() == 1) print *, x
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

#[test] fn co_sum_array() {
    let out = run_prints(r#"
program test
    integer :: a(3)
    a = this_image()
    call co_sum(a)
    print *, a(1)
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

#[test] fn co_sum_result_image() {
    let out = run_prints(r#"
program test
    real :: x = 1.0
    call co_sum(x, result_image=1)
    if (this_image() == 1) print *, x
end program test
"#);
    assert_eq!(out, vec!["1"]);
}

#[test] fn co_sum_with_stat() {
    let out = run_prints(r#"
program test
    integer :: x = 5, stat
    call co_sum(x, stat=stat)
    print *, x
end program test
"#);
    assert_eq!(out, vec!["5"]);
}

// ── CO_MAX collective ─────────────────────────────────────────

#[test] fn co_max_scalar() {
    compile_ok(r#"
program test
    integer :: x
    x = this_image() * 10
    call co_max(x)
    print *, x
end program test
"#);
}

#[test] fn co_max_real() {
    compile_ok(r#"
program test
    real :: r = 3.14 * this_image()
    call co_max(r, result_image=1)
    if (this_image() == 1) print *, r
end program test
"#);
}

// ── CO_MIN collective ─────────────────────────────────────────

#[test] fn co_min_scalar() {
    compile_ok(r#"
program test
    integer :: x = 100 - this_image()
    call co_min(x)
    print *, x
end program test
"#);
}

// ── CO_BROADCAST collective ───────────────────────────────────

#[test] fn co_broadcast_integer() {
    compile_ok(r#"
program test
    integer :: x
    if (this_image() == 1) x = 42
    call co_broadcast(x, source_image=1)
    print *, x
end program test
"#);
}

#[test] fn co_broadcast_character() {
    compile_ok(r#"
program test
    character(len=20) :: msg
    if (this_image() == 1) msg = 'hello from 1'
    call co_broadcast(msg, source_image=1)
    print *, trim(msg)
end program test
"#);
}

#[test] fn co_broadcast_array() {
    compile_ok(r#"
program test
    integer :: a(5)
    if (this_image() == 1) a = [1, 2, 3, 4, 5]
    call co_broadcast(a, source_image=1)
    print *, a(3)
end program test
"#);
}

// ── CO_REDUCE collective ──────────────────────────────────────

#[test] fn co_reduce_sum() {
    compile_ok(r#"
program test
    integer :: x = this_image()
    call co_reduce(x, operator(+))
    if (this_image() == 1) print *, x
end program test
"#);
}

#[test] fn co_reduce_user_op() {
    compile_ok(r#"
program test
    integer :: x = this_image()
    call co_reduce(x, my_add, result_image=1)
    if (this_image() == 1) print *, x
contains
    pure function my_add(a, b) result(c)
        integer, intent(in) :: a, b
        integer :: c
        c = a + b
    end function my_add
end program test
"#);
}

// ── Remote coarray access (bracket indexing) ──────────────────

#[test] fn coarray_remote_read() {
    compile_ok(r#"
program test
    integer :: x[*]
    x = this_image() * 10
    sync all
    if (this_image() == 1 .and. num_images() >= 2) then
        print *, x[2]
    end if
end program test
"#);
}

#[test] fn coarray_remote_write() {
    compile_ok(r#"
program test
    integer :: shared[*]
    shared = 0
    sync all
    if (this_image() == 1) then
        shared[1] = 99
    end if
    sync all
    if (this_image() == 1) print *, shared
end program test
"#);
}

#[test] fn coarray_array_remote_element() {
    compile_ok(r#"
program test
    integer :: a(5)[*]
    a = this_image()
    sync all
    if (this_image() == 1 .and. num_images() >= 2) then
        print *, a(3)[2]
    end if
end program test
"#);
}

// ── Allocatable coarrays ──────────────────────────────────────

#[test] fn allocatable_coarray_basic() {
    compile_ok(r#"
program test
    integer, allocatable :: x[:]
    allocate(x[*])
    x = this_image()
    print *, x
    deallocate(x)
end program test
"#);
}

#[test] fn allocatable_coarray_array() {
    compile_ok(r#"
program test
    real, allocatable :: a(:)[:]
    allocate(a(10)[*])
    a = real(this_image())
    print *, a(1)
    deallocate(a)
end program test
"#);
}

// ── EVENT POST / WAIT ─────────────────────────────────────────

#[test] fn event_post_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    if (this_image() == 1) then
        event post(ev[2])
    else if (this_image() == 2) then
        event wait(ev)
        print *, 'event received'
    end if
end program test
"#);
}

#[test] fn event_query_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: count
    event query(ev, count)
    print *, count
end program test
"#);
}

#[test] fn event_wait_until_count() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(event_type) :: ev[*]
    if (this_image() == 1) then
        event wait(ev, until_count=3)
        print *, 'got 3 events'
    end if
end program test
"#);
}

// ── CRITICAL construct with coarrays ─────────────────────────

#[test] fn critical_coarray_update() {
    compile_ok(r#"
program test
    integer :: counter[*]
    counter = 0
    sync all
    critical
        counter[1] = counter[1] + 1
    end critical
    sync all
    if (this_image() == 1) print *, counter
end program test
"#);
}

// ── LOCK / UNLOCK ─────────────────────────────────────────────

#[test] fn lock_unlock_coarray() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    lock(lk[1])
    print *, 'locked'
    unlock(lk[1])
    print *, 'unlocked'
end program test
"#);
}

#[test] fn lock_with_stat() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk, stat=stat, acquired_lock=.true.)
    if (stat == 0) then
        print *, 'locked ok'
        unlock(lk)
    end if
end program test
"#);
}

// ── Coarray in derived type ───────────────────────────────────

#[test] fn coarray_derived_type_member() {
    compile_ok(r#"
program test
    type :: Shared
        integer :: value
    end type Shared
    type(Shared) :: obj[*]
    obj%value = this_image()
    sync all
    print *, obj%value
end program test
"#);
}

// ── FAILED_IMAGES / STOPPED_IMAGES (F2018) ───────────────────

#[test] fn failed_images_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer, allocatable :: fi(:)
    fi = failed_images()
    print *, size(fi)
end program test
"#);
}

#[test] fn stopped_images_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer, allocatable :: si(:)
    si = stopped_images()
    print *, size(si)
end program test
"#);
}

// ── IMAGE_STATUS intrinsic (F2018) ────────────────────────────

#[test] fn image_status_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer :: status
    status = image_status(1)
    print *, status
end program test
"#);
}

// ── Coarray bounds — cobounds ─────────────────────────────────

#[test] fn lcobound_ucobound() {
    compile_ok(r#"
program test
    integer :: x[2:4, *]
    print *, lcobound(x, 1)
    print *, ucobound(x, 1)
end program test
"#);
}

#[test] fn lcobound_dim() {
    compile_ok(r#"
program test
    integer :: x[3:*]
    print *, lcobound(x)
end program test
"#);
}

// ── GET_TEAM / CHANGE TEAM ────────────────────────────────────

#[test] fn form_team_and_change() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(team_type) :: odd_even
    integer :: color
    color = mod(this_image(), 2) + 1
    call form_team(color, odd_even)
    change team (odd_even)
        print *, this_image(), 'in subteam', team_number()
    end team
    sync all
    if (this_image() == 1) print *, 'done'
end program test
"#);
}

#[test] fn team_number_in_team() {
    compile_ok(r#"
program test
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        print *, team_number()
    end team
end program test
"#);
}

// ── Parallel reduction pattern ────────────────────────────────

#[test] fn parallel_sum_pattern() {
    compile_ok(r#"
program test
    integer :: x[*], total
    x = this_image()
    sync all
    if (this_image() == 1) then
        total = 0
        integer :: i
        do i = 1, num_images()
            total = total + x[i]
        end do
        print *, total
    end if
end program test
"#);
}

#[test] fn parallel_broadcast_pattern() {
    compile_ok(r#"
program test
    integer :: seed[*]
    if (this_image() == 1) seed = 42
    call co_broadcast(seed, source_image=1)
    print *, seed
end program test
"#);
}

// ── ATOMIC operations (F2018) ─────────────────────────────────

#[test] fn atomic_define_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    call atomic_define(counter, 0)
    print *, counter
end program test
"#);
}

#[test] fn atomic_ref_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    integer :: val
    call atomic_define(counter, 7)
    call atomic_ref(val, counter)
    print *, val
end program test
"#);
}

#[test] fn atomic_add_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(atomic_int_kind) :: n[*]
    call atomic_define(n, 0)
    call atomic_add(n, 1)
    print *, n
end program test
"#);
}

#[test] fn atomic_cas_basic() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(atomic_int_kind) :: x[*]
    integer :: old
    call atomic_define(x, 10)
    call atomic_cas(x, old, 10, 20)
    print *, x
end program test
"#);
}

#[test] fn atomic_fetch_add() {
    compile_ok(r#"
program test
    use iso_fortran_env
    integer(atomic_int_kind) :: counter[*]
    integer :: prev
    call atomic_define(counter, 5)
    call atomic_fetch_add(counter, 3, prev)
    print *, prev
    print *, counter
end program test
"#);
}

#[test] fn atomic_logical_ops() {
    compile_ok(r#"
program test
    use iso_fortran_env
    logical(atomic_logical_kind) :: flag[*]
    call atomic_define(flag, .false.)
    call atomic_or(flag, .true.)
    print *, flag
end program test
"#);
}
