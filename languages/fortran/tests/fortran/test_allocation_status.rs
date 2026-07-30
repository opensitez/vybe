use super::helpers::{compile_ok, run_prints};
macro_rules! c {
    ($n:ident,$s:expr) => {
        #[test]
        fn $n() {
            compile_ok($s);
        }
    };
}
c!(
    allocation_status_01,
    "program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3), stat=st)
end program p
"
);
c!(
    allocation_status_02,
    "program p
integer, allocatable :: a(:)
integer :: st
allocate(a(3))
deallocate(a, stat=st)
end program p
"
);
c!(
    allocation_status_03,
    "program p
integer, allocatable :: x
integer :: st
allocate(x, stat=st)
end program p
"
);
c!(
    allocation_status_04,
    "program p
character(len=20) :: msg
integer :: st
integer, allocatable :: a(:)
allocate(a(3), stat=st, errmsg=msg)
end program p
"
);
c!(
    allocation_status_05,
    "program p
character(len=20) :: msg
integer :: st
integer, allocatable :: a(:)
allocate(a(3))
deallocate(a, stat=st, errmsg=msg)
end program p
"
);
c!(
    allocation_status_06,
    "program p
integer, pointer :: p(:)
integer :: st
allocate(p(3), stat=st)
end program p
"
);
c!(
    allocation_status_07,
    "program p
class(*), allocatable :: x
integer :: st
allocate(integer :: x, stat=st)
end program p
"
);
c!(
    allocation_status_08,
    "program p
logical, allocatable :: a(:)
integer :: st
allocate(a(2), stat=st)
end program p
"
);
c!(
    allocation_status_09,
    "program p
complex, allocatable :: a(:)
integer :: st
allocate(a(2), stat=st)
end program p
"
);
c!(
    allocation_status_10,
    "program p
real, allocatable :: a(:,:)
integer :: st
allocate(a(2,2), stat=st)
end program p
"
);

#[test]
fn allocation_status_records_zero_for_successful_alloc() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a(:)\n\
integer :: st\n\
allocate(a(3), stat=st)\n\
if (st == 0) then\n\
  print *, 'ok'\n\
else\n\
  print *, 'bad'\n\
end if\n\
end program t\n"
        ),
        vec!["ok"]
    );
}

#[test]
fn allocation_status_deallocate_unallocated_reports_nonzero_in_stat() {
    assert_eq!(
        run_prints(
            "program t\n\
integer, allocatable :: a(:)\n\
integer :: st\n\
deallocate(a, stat=st)\n\
if (st == 0) then\n\
  print *, 'ok'\n\
else\n\
  print *, 'bad'\n\
end if\n\
end program t\n"
        ),
        vec!["bad"]
    );
}
