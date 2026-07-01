use super::helpers::compile_ok;
macro_rules! c { ($n:ident,$s:expr)=>{ #[test] fn $n(){ compile_ok($s); } }; }
c!(co_img_ctl_01,"program p
integer :: x[*]
sync all
end program p
");
c!(co_teams_02,"program p
form team (1, team1)
end program p
");
c!(co_events_03,"program p
type(event_type) :: ev[*]
end program p
");
c!(co_locks_04,"program p
type(lock_type) :: lk[*]
end program p
");
c!(co_atomic_05,"program p
integer(atomic_int_kind) :: x[*]
end program p
");
c!(co_coidx_06,"program p
integer :: x[*]
print *, x[1]
end program p
");
c!(co_failed_07,"program p
integer :: i
i = failed_images()
print *, i
end program p
");
c!(co_stopped_08,"program p
integer :: i
i = stopped_images()
print *, i
end program p
");
c!(co_collective_09,"program p
integer :: x[*]
call co_sum(x)
end program p
");
c!(co_status_10,"program p
integer :: s
s = image_status(1)
print *, s
end program p
");
c!(co_sync_mem_11,"program p
sync memory
end program p
");
c!(co_sync_images_12,"program p
sync images(*)
end program p
");
c!(co_sync_team_13,"program p
sync team(1)
end program p
");
c!(co_this_image_14,"program p
print *, this_image()
end program p
");
c!(co_num_images_15,"program p
print *, num_images()
end program p
");
c!(co_form_team_16,"program p
integer :: team1
form team(1, team1)
end program p
");
c!(co_change_team_17,"program p
change team (team=1)
end team
end program p
");
c!(co_end_team_18,"program p
change team (team=1)
 print *,1
end team
end program p
");
c!(co_atomic_define_19,"program p
use iso_fortran_env
integer(atomic_int_kind) :: x[*]
call atomic_define(x,1)
end program p
");
c!(co_atomic_ref_20,"program p
use iso_fortran_env
integer(atomic_int_kind) :: x[*], y
call atomic_ref(y,x)
end program p
");