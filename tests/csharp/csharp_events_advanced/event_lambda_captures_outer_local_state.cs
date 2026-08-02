// vybe-test: csharp/csharp_events_advanced/event_lambda_captures_outer_local_state
// origin: languages/csharp/tests/csharp/test_csharp_events_advanced.rs

using System; class Alarm { public event Action Triggered; public void Fire() { Triggered(); } } int count = 0; var alarm = new Alarm(); alarm.Triggered += () => { count++; Console.WriteLine(count); }; alarm.Fire(); alarm.Fire();
