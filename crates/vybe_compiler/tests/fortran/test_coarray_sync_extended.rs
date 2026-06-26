//! Extended coarray sync / team / event coverage: barrier variants, image
//! subsets, lock/unlock, critical, and event post/wait forms not in
//! `test_coarrays.rs`. Distinct only — no duplicate scenarios.

use super::helpers::compile_ok;

fortran_cases! {
    // ── SYNC ALL (extended) ─────────────────────────────────────────

    sync_all_double_barrier_prints_once => {
        "program t\ninteger :: step[*]\nstep = this_image()\nsync all\nsync all\nif (this_image() == 1) print *, step\nend program t\n",
        ["1"]
    };

    sync_all_stat_errmsg_reports_ok => {
        "program t\ninteger :: stat\ncharacter(len=80) :: errmsg\nsync all (stat=stat, errmsg=errmsg)\nif (stat == 0) print *, 'barrier ok'\nend program t\n",
        ["barrier ok"]
    };

    sync_all_after_remote_read_pattern => {
        "program t\ninteger :: buf[*]\nbuf = this_image() * 3\nsync all\nif (this_image() == 1) print *, buf\nend program t\n",
        ["3"]
    };

    // ── SYNC IMAGES (extended) ──────────────────────────────────────

    sync_images_star_stat_zero => {
        "program t\ninteger :: stat\nsync images (*, stat=stat)\nprint *, stat\nend program t\n",
        ["0"]
    };

    sync_images_int_array_target => {
        "program t\ninteger :: peers(1)\npeers(1) = this_image()\nsync images (peers)\nprint *, 'peers ok'\nend program t\n",
        ["peers ok"]
    };

    sync_images_bracket_list_literal => {
        "program t\nsync images ([1])\nprint *, 'self sync'\nend program t\n",
        ["self sync"]
    };

    // ── SYNC MEMORY (extended) ──────────────────────────────────────

    sync_memory_stat_clause => {
        "program t\ninteger :: stat, x[*]\nx = 0\nsync memory (stat=stat)\nx = 1\nprint *, x, stat\nend program t\n",
        ["1 0"]
    };

    // ── SYNC TEAM (extended) ────────────────────────────────────────

    sync_team_on_initial_team => {
        "program t\nuse iso_fortran_env\ntype(team_type) :: init\ntype(event_type) :: ev[*]\ninit = get_team(initial_team)\nsync team (init)\nprint *, team_number(init)\nend program t\n",
        ["-1"]
    };

    // ── EVENT post / wait / query (extended) ────────────────────────

    event_post_then_query_count => {
        "program t\nuse iso_fortran_env\ntype(event_type) :: ev[*]\ninteger :: count\nevent post(ev)\nevent query(ev, count)\nprint *, count\nend program t\n",
        ["1"]
    };

    event_wait_until_count_one => {
        "program t\nuse iso_fortran_env\ntype(event_type) :: ev[*]\nevent post(ev)\nevent wait(ev, until_count=1)\nprint *, 'ready'\nend program t\n",
        ["ready"]
    };

    // ── CRITICAL (extended) ─────────────────────────────────────────

    critical_multi_statement_updates_sum => {
        "program t\ninteger :: a = 0, b = 0\ncritical\na = a + 2\nb = b + a\nend critical\nprint *, a, b\nend program t\n",
        ["2 2"]
    };

    critical_coarray_real_accumulator => {
        "program t\nreal :: total[*]\ntotal = 0.0\nsync all\ncritical\ntotal[1] = total[1] + real(this_image())\nend critical\nsync all\nif (this_image() == 1) print *, total\nend program t\n",
        ["1"]
    };

    // ── LOCK / UNLOCK (extended) ────────────────────────────────────

    lock_unlock_local_image_prints => {
        "program t\nuse iso_fortran_env\ninteger(lock_type) :: lk[*]\nlock(lk)\nprint *, 'held'\nunlock(lk)\nprint *, 'free'\nend program t\n",
        ["held", "free"]
    };
}

// ── SYNC ALL: conditional barrier (compile-only) ─────────────────────

#[test]
fn sync_all_only_on_leader_image() {
    compile_ok(
        r#"
program t
    integer :: flag[*]
    flag = 0
    if (this_image() == 1) sync all
    flag = this_image()
    print *, flag
end program t
"#,
    );
}

// ── SYNC IMAGES: image-index expression (compile-only) ───────────────

#[test]
fn sync_images_this_image_list() {
    compile_ok(
        r#"
program t
    sync images ([this_image()])
    print *, 'done'
end program t
"#,
    );
}

// ── SYNC TEAM: after form_team (compile-only) ────────────────────────

#[test]
fn sync_team_on_formed_subteam() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    type(team_type) :: sub
    call form_team(mod(this_image(), 2) + 1, sub)
    sync team (sub)
    print *, 'subteam synced'
end program t
"#,
    );
}

// ── CHANGE TEAM with inner sync all (compile-only) ─────────────────

#[test]
fn change_team_inner_sync_all() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    type(team_type) :: t
    call form_team(1, t)
    change team (t)
        sync all
        print *, team_number()
    end team
end program t
"#,
    );
}

// ── EVENT: stat / errmsg clauses (compile-only) ──────────────────────

#[test]
fn event_post_with_stat() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: stat
    event post(ev, stat=stat)
    print *, stat
end program t
"#,
    );
}

#[test]
fn event_wait_stat_errmsg() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    type(event_type) :: ev[*]
    integer :: stat
    character(len=80) :: errmsg
    event post(ev)
    event wait(ev, stat=stat, errmsg=errmsg)
    print *, 'waited'
end program t
"#,
    );
}

// ── NOTIFY post / wait (F2018, compile-only) ───────────────────────

#[test]
fn notify_post_wait_pair() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    type(event_type) :: note[*]
    notify post(note[1])
    notify wait(note)
    print *, 'notify done'
end program t
"#,
    );
}

// ── CRITICAL: stat clause and named construct (compile-only) ─────────

#[test]
fn critical_with_stat_clause() {
    compile_ok(
        r#"
program t
    integer :: n = 0
    integer :: stat
    critical (stat=stat)
        n = n + 1
    end critical
    print *, n, stat
end program t
"#,
    );
}

#[test]
fn critical_named_construct_label() {
    compile_ok(
        r#"
program t
    integer :: slot[*]
    slot = 0
    sync all
    guard: critical
        slot[1] = slot[1] + this_image()
    end critical guard
    sync all
    if (this_image() == 1) print *, slot
end program t
"#,
    );
}

// ── LOCK / UNLOCK: stat, errmsg, acquired_lock (compile-only) ──────

#[test]
fn unlock_with_stat() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk)
    unlock(lk, stat=stat)
    print *, stat
end program t
"#,
    );
}

#[test]
fn lock_errmsg_when_busy() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    character(len=80) :: errmsg
    lock(lk, stat=stat, errmsg=errmsg, acquired_lock=.true.)
    if (stat == 0) unlock(lk)
    print *, trim(errmsg)
end program t
"#,
    );
}

#[test]
fn lock_acquired_false_nonblocking() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    integer(lock_type) :: lk[*]
    integer :: stat
    lock(lk, stat=stat, acquired_lock=.false.)
    print *, stat
end program t
"#,
    );
}

#[test]
fn lock_unlock_inside_critical() {
    compile_ok(
        r#"
program t
    use iso_fortran_env
    integer :: tally[*]
    integer(lock_type) :: lk[*]
    tally = 0
    sync all
    critical
        lock(lk[1])
        tally[1] = tally[1] + 1
        unlock(lk[1])
    end critical
    sync all
    if (this_image() == 1) print *, tally
end program t
"#,
    );
}
