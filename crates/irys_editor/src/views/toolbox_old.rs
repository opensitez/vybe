use crate::message::Message;
use iced::widget::{button, column, container, text};
use iced::{Element, Length};
use irys_forms::ControlType;

pub fn toolbox() -> Element<'static, Message> {
    let controls = vec![
        ("Button", ControlType::Button),
        ("Label", ControlType::Label),
        ("TextBox", ControlType::TextBox),
        ("CheckBox", ControlType::CheckBox),
        ("RadioButton", ControlType::RadioButton),
    ];

    let mut col = column![
        text("Toolbox").size(18),
        container(text("").size(1))
            .width(Length::Fill)
            .height(Length::Fixed(1.0))
            .style(|_theme| {
                container::Style {
                    background: Some(iced::Background::Color(iced::Color::from_rgb(0.7, 0.7, 0.7))),
                    ..Default::default()
                }
            })
    ]
    .spacing(10)
    .padding(10);

    col = col.push(
        text("Controls")
            .size(12)
            .style(|_theme| {
                text::Style {
                    color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                }
            })
    );

    for (label, control_type) in controls {
        col = col.push(
            button(
                text(label)
                    .size(13)
            )
            .width(Length::Fill)
            .padding(8)
            .on_press(Message::SelectTool(control_type))
            .style(|theme, status| {
                button::Style {
                    background: Some(iced::Background::Color(
                        match status {
                            button::Status::Hovered => iced::Color::from_rgb(0.85, 0.85, 0.85),
                            button::Status::Pressed => iced::Color::from_rgb(0.75, 0.75, 0.75),
                            _ => iced::Color::from_rgb(0.95, 0.95, 0.95),
                        }
                    )),
                    border: iced::Border {
                        color: iced::Color::from_rgb(0.7, 0.7, 0.7),
                        width: 1.0,
                        radius: 3.0.into(),
                    },
                    text_color: iced::Color::BLACK,
                    ..Default::default()
                }
            })
        );
    }

    col = col.push(text("").size(10));
    col = col.push(
        column![
            text("Instructions:").size(11),
            text("1. Select a control").size(10),
            text("2. Click on canvas").size(10),
            text("3. Edit properties").size(10),
        ]
        .spacing(3)
        .padding(5)
        .style(|_theme| {
            text::Style {
                color: Some(iced::Color::from_rgb(0.5, 0.5, 0.5)),
            }
        })
    );

    container(col)
        .width(Length::Fixed(160.0))
        .height(Length::Fill)
        .style(|_theme| {
            container::Style {
                background: Some(iced::Background::Color(iced::Color::from_rgb(0.98, 0.98, 0.98))),
                border: iced::Border {
                    color: iced::Color::from_rgb(0.85, 0.85, 0.85),
                    width: 1.0,
                    ..Default::default()
                },
                ..Default::default()
            }
        })
        .into()
}
