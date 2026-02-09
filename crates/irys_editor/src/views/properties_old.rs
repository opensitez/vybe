use crate::message::Message;
use crate::state::EditorState;
use iced::widget::{button, checkbox, column, container, row, scrollable, text, text_input};
use iced::{Element, Length};

pub fn properties_panel(state: &EditorState) -> Element<'static, Message> {
    let mut col = column![
        text("Properties").size(18),
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

    if let Some(control) = state.get_selected_control() {
        col = col.push(
            column![
                text("Name").size(11).style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                    }
                }),
                text(control.name.clone()).size(13),
            ]
            .spacing(3)
        );

        col = col.push(
            column![
                text("Type").size(11).style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                    }
                }),
                text(control.control_type.as_str()).size(13),
            ]
            .spacing(3)
        );

        // Caption property
        if let Some(caption) = control.get_caption() {
            col = col.push(
                column![
                    text("Caption").size(11).style(|_theme| {
                        text::Style {
                            color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                        }
                    }),
                    text_input("", caption)
                        .on_input(|s| Message::PropertyChanged("Caption".to_string(), s))
                        .padding(5)
                        .size(13)
                ]
                .spacing(3)
            );
        }

        // Text property
        if let Some(txt) = control.get_text() {
            col = col.push(
                column![
                    text("Text").size(11).style(|_theme| {
                        text::Style {
                            color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                        }
                    }),
                    text_input("", txt)
                        .on_input(|s| Message::PropertyChanged("Text".to_string(), s))
                        .padding(5)
                        .size(13)
                ]
                .spacing(3)
            );
        }

        // Enabled property
        col = col.push(
            row![
                text("Enabled").size(13),
                checkbox("", control.is_enabled())
                    .on_toggle(|val| Message::PropertyChanged("Enabled".to_string(), val.to_string()))
            ]
            .spacing(10)
        );

        // Visible property
        col = col.push(
            row![
                text("Visible").size(13),
                checkbox("", control.is_visible())
                    .on_toggle(|val| Message::PropertyChanged("Visible".to_string(), val.to_string()))
            ]
            .spacing(10)
        );

        // Layout section
        col = col.push(
            column![
                text("").size(5),
                text("Layout").size(14),
                row![
                    column![
                        text("X").size(11).style(|_theme| {
                            text::Style {
                                color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                            }
                        }),
                        text_input("", &control.bounds.x.to_string())
                            .on_input(|s| Message::PropertyChanged("X".to_string(), s))
                            .padding(5)
                            .size(13)
                            .width(Length::Fixed(70.0))
                    ]
                    .spacing(3),
                    column![
                        text("Y").size(11).style(|_theme| {
                            text::Style {
                                color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                            }
                        }),
                        text_input("", &control.bounds.y.to_string())
                            .on_input(|s| Message::PropertyChanged("Y".to_string(), s))
                            .padding(5)
                            .size(13)
                            .width(Length::Fixed(70.0))
                    ]
                    .spacing(3),
                ]
                .spacing(10),
                row![
                    column![
                        text("Width").size(11).style(|_theme| {
                            text::Style {
                                color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                            }
                        }),
                        text_input("", &control.bounds.width.to_string())
                            .on_input(|s| Message::PropertyChanged("Width".to_string(), s))
                            .padding(5)
                            .size(13)
                            .width(Length::Fixed(70.0))
                    ]
                    .spacing(3),
                    column![
                        text("Height").size(11).style(|_theme| {
                            text::Style {
                                color: Some(iced::Color::from_rgb(0.4, 0.4, 0.4)),
                            }
                        }),
                        text_input("", &control.bounds.height.to_string())
                            .on_input(|s| Message::PropertyChanged("Height".to_string(), s))
                            .padding(5)
                            .size(13)
                            .width(Length::Fixed(70.0))
                    ]
                    .spacing(3),
                ]
                .spacing(10),
            ]
            .spacing(5)
        );

        // Delete button
        col = col.push(
            column![
                text("").size(10),
                button(text("Delete Control").size(12))
                    .on_press(Message::DeleteControl)
                    .width(Length::Fill)
            ]
        );

    } else {
        col = col.push(
            text("No control selected")
                .size(13)
                .style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.5, 0.5, 0.5)),
                    }
                })
        );
        col = col.push(text("").size(5));
        col = col.push(
            text("Click on a control to select it")
                .size(12)
                .style(|_theme| {
                    text::Style {
                        color: Some(iced::Color::from_rgb(0.5, 0.5, 0.5)),
                    }
                })
        );
    }

    container(
        scrollable(col)
    )
    .width(Length::Fixed(220.0))
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
