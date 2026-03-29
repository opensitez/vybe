mod app;
mod state;
mod panels;

fn main() {
    let cli_path = std::env::args().nth(1).map(std::path::PathBuf::from);

    let options = eframe::NativeOptions {
        viewport: egui::ViewportBuilder::default()
            .with_title("vybe Basic IDE")
            .with_inner_size([1280.0, 800.0]),
        ..Default::default()
    };

    eframe::run_native(
        "vybe Basic IDE",
        options,
        Box::new(move |cc| Ok(Box::new(app::VybeApp::new(cc, cli_path)))),
    ).unwrap();
}
