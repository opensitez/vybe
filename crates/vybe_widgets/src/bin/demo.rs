use tiny_skia::{Pixmap, Paint, Rect};
use cosmic_text::{FontSystem, SwashCache, Color as CColor};

use vybe_widgets::{Button, Checkbox, TextInput, Radio, Slider, Dropdown, TreeView};

fn main() {
    let width = 800;
    let height = 400;
    let mut pixmap = Pixmap::new(width, height).expect("create pixmap");

    // Font system for text rendering
    let mut fs = FontSystem::new();
    let mut sc = SwashCache::new();

    // Background
    let mut bg = Paint::default();
    bg.set_color_rgba8(250, 250, 250, 255);
    pixmap.fill_rect(Rect::from_xywh(0.0, 0.0, width as f32, height as f32).unwrap(), &bg, tiny_skia::Transform::identity(), None);

    // Checkbox
    let mut checkbox = Checkbox::new("Accept terms");
    checkbox.check_state = vybe_widgets::CheckState::Checked;
    checkbox.paint(&mut pixmap, 20.0, 20.0, 1.0);
    TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, "Accept terms", 46.0, 18.0, CColor::rgb(51,51,51), 1.0);

    // Button
    let mut button = Button::new("Click me");
    button.paint(&mut pixmap, 20.0, 60.0, 1.0);
    TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, "Click me", 30.0, 64.0, CColor::rgb(51,51,51), 1.0);

    // Text input
    let mut ti = TextInput::new().with_placeholder("Enter your name");
    ti.paint_border(&mut pixmap, 20.0, 110.0, 1.0);
    let disp = ti.display_text();
    let col = if ti.is_placeholder() { CColor::rgb(150,150,150) } else { CColor::rgb(0,0,0) };
    TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, &disp, 26.0, 112.0, col, 1.0);

    // Radio
    let mut radio = Radio::new("Option A");
    radio.selected = true;
    radio.paint(&mut pixmap, 20.0, 160.0, 1.0);
    TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, "Option A", 46.0, 158.0, CColor::rgb(51,51,51), 1.0);

    // Slider
    let mut slider = Slider::new(0.0, 100.0, 42.0);
    slider.paint(&mut pixmap, 20.0, 200.0, 1.0);
    TreeView::draw_text_static_internal(&mut pixmap, &mut fs, &mut sc, &format!("Value: {:.0}", slider.actual_value()), 240.0, 200.0, CColor::rgb(51,51,51), 1.0);

    // Dropdown
    let items = vec!["First".to_string(), "Second".to_string(), "Third".to_string(), "Fourth".to_string()];
    let dd = Dropdown::new(items, 1, 1.0, None);
    dd.render_list(&mut pixmap, &mut fs, &mut sc, 20.0, 250.0, (255,255,255,255), (180,180,180,255), (220,220,220,255), (240,240,240,255), CColor::rgb(0,0,0), CColor::rgb(120,120,120));

    // Save output image
    pixmap.save_png("widgets_demo.png").expect("save png");
    println!("Wrote widgets_demo.png ({}x{})", width, height);
}
