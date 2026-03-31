use std::{collections::HashSet, thread, time::Duration};

use rand::RngExt;
use rust_drission::{Browser, BrowserConfig, CdpError, Page};
use rust_xlsxwriter::{Workbook, XlsxError};

#[derive(Debug)]
pub struct XhsNoteDetail {
    pub title: String,
    pub imgs: HashSet<String>,
    pub bloger: String,
    pub content: String,
    pub tags: Vec<String>,
    pub publish_time: String,
    pub note_link: String,
    pub like_count: u32,
    pub collect_count: u32,
    pub comment_count: u32,
    pub comment_details: Vec<String>,
}

fn export_to_xlsx(path: &str, details: &[XhsNoteDetail]) -> Result<(), XlsxError> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let max_img_count = details.iter().map(|d| d.imgs.len()).max().unwrap_or(0);

    let mut headers = vec![
        "标题".to_string(),
        "博主".to_string(),
        "发布时间".to_string(),
        "笔记链接".to_string(),
        "点赞数".to_string(),
        "收藏数".to_string(),
        "评论数".to_string(),
        "标签".to_string(),
        "评论详情".to_string(),
        "内容".to_string(),
    ];
    for idx in 0..max_img_count {
        headers.push(format!("图片链接{}", idx + 1));
    }

    for (col, header) in headers.iter().enumerate() {
        worksheet.write_string(0, col as u16, header)?;
    }

    for (idx, detail) in details.iter().enumerate() {
        let row = (idx + 1) as u32;
        let mut imgs = detail.imgs.iter().cloned().collect::<Vec<_>>();
        imgs.sort();

        worksheet.write_string(row, 0, &detail.title)?;
        worksheet.write_string(row, 1, &detail.bloger)?;
        worksheet.write_string(row, 2, &detail.publish_time)?;
        worksheet.write_url_with_text(
            row,
            3,
            detail.note_link.as_str(),
            detail.note_link.as_str(),
        )?;
        worksheet.write_number(row, 4, detail.like_count as f64)?;
        worksheet.write_number(row, 5, detail.collect_count as f64)?;
        worksheet.write_number(row, 6, detail.comment_count as f64)?;
        worksheet.write_string(row, 7, detail.tags.join(", "))?;
        worksheet.write_string(row, 8, detail.comment_details.join("\n"))?;
        worksheet.write_string(row, 9, &detail.content)?;

        for (img_idx, img) in imgs.iter().enumerate() {
            let img_col = (10 + img_idx) as u16;
            worksheet.write_url_with_text(row, img_col, img.as_str(), img.as_str())?;
        }
    }

    worksheet.set_column_width(0, 30)?;
    worksheet.set_column_width(1, 16)?;
    worksheet.set_column_width(2, 16)?;
    worksheet.set_column_width(3, 48)?;
    worksheet.set_column_width(4, 10)?;
    worksheet.set_column_width(5, 10)?;
    worksheet.set_column_width(6, 10)?;
    worksheet.set_column_width(7, 24)?;
    worksheet.set_column_width(8, 48)?;
    worksheet.set_column_width(9, 64)?;
    for img_idx in 0..max_img_count {
        worksheet.set_column_width((10 + img_idx) as u16, 16)?;
    }

    workbook.save(path)?;
    Ok(())
}

fn handle_detail(page: &Page) -> Result<XhsNoteDetail, CdpError> {
    let mut note_detail = XhsNoteDetail {
        title: String::new(),
        imgs: HashSet::new(),
        bloger: String::new(),
        content: String::new(),
        tags: Vec::new(),
        publish_time: String::new(),
        note_link: page.url()?,
        like_count: 0,
        collect_count: 0,
        comment_count: 0,
        comment_details: Vec::new(),
    };
    // 获取笔记标题
    let title = page.element("#detail-title")?;
    if let Some(title) = title {
        note_detail.title = title.text()?;
    }
    // 获取笔记图片
    let img_eles = page.elements(".img-container img")?;
    for img_ele in img_eles {
        let src = img_ele.attr("src")?;
        note_detail.imgs.insert(src);
    }

    // 获取博主名称
    let username_ele = page.elements(".username")?;
    if let Some(username_ele) = username_ele.first() {
        note_detail.bloger = username_ele.text()?;
    }

    // 获取笔记详情
    let desc = page.element("#detail-desc")?;
    if let Some(desc) = desc {
        note_detail.content = desc.text()?;
    }

    // 获取Tag
    let tags_eles = page.elements(".note-text .tag")?;
    for tag_ele in tags_eles {
        let tag_text = tag_ele.text()?;
        note_detail.tags.push(tag_text);
    }

    // 获取发布时间
    let publish_time_ele = page.element(".bottom-container .date")?;
    if let Some(publish_time_ele) = publish_time_ele {
        note_detail.publish_time = publish_time_ele.text()?;
    }

    // 获取点赞、收藏、评论数
    let like_ele = page.element(".interactions.engage-bar .like-wrapper .count")?;
    if let Some(like_ele) = like_ele {
        note_detail.like_count = like_ele.text()?.parse().unwrap_or(0);
    }

    let collect_ele = page.element(".interactions.engage-bar .collect-wrapper  .count")?;
    if let Some(collect_ele) = collect_ele {
        note_detail.collect_count = collect_ele.text()?.parse().unwrap_or(0);
    }
    let comment_ele = page.element(".interactions.engage-bar .chat-wrapper  .count")?;
    if let Some(comment_ele) = comment_ele {
        note_detail.comment_count = comment_ele.text()?.parse().unwrap_or(0);
    }

    // 获取评论详情
    // document.querySelectorAll(".parent-comment")[2].querySelectorAll(".note-text")[0].innerText
    let commen_eles = page.elements(".parent-comment")?;
    for commen_ele in commen_eles {
        let note_text_ele = commen_ele.element(".note-text")?;
        if let Some(note_text_ele) = note_text_ele {
            note_detail.comment_details.push(note_text_ele.text()?);
        }
    }

    return Ok(note_detail);
}

fn handle_scroll(page: &Page) -> Result<(), CdpError> {
    let scroll_js = r#"
  window.scrollTo({
    top: document.body.scrollHeight,
    behavior: 'smooth'
  });
    "#;
    page.run_js(scroll_js)?;
    Ok(())
}

// 随机睡眠 单位ms
fn random_random_sleep(min: u64, max: u64) {
    let random_sleep_time = rand::rng().random_range(min..max);
    thread::sleep(Duration::from_millis(random_sleep_time));
}

fn main() -> Result<(), CdpError> {
    // 获取当前 exe 路径
    let exe_path = std::env::current_exe().expect("Failed to get exe path");

    // 获取 exe 所在目录
    let exe_dir = exe_path.parent().expect("Failed to get exe dir");

    // 构造 .env 路径
    let env_path = exe_dir.join(".env");

    println!("env_path: {}", env_path.display());

    // 手动加载 .env
    dotenvy::from_path(&env_path).ok();
    let port = std::env::var("PORT").unwrap_or_else(|_| "9225".into());
    let port = port.parse::<u16>().unwrap();
    let keyword = std::env::var("KEYWORD").unwrap_or_else(|_| "erp".into());
    let sort_type = std::env::var("TYPE").unwrap_or_else(|_| "最多点赞".into());
    let browser_path = std::env::var("BROWSER_PATH").unwrap_or_else(|_| "".into());
    let user_data_dir =
        std::env::var("USER_DATA_DIR").unwrap_or_else(|_| r"E:\tmp\UserData\XHS".into());
    let like_limit = std::env::var("LIKE_LOWER_LIMIT")
        .unwrap_or_else(|_| "10".into())
        .parse::<u32>()
        .unwrap();
    let comment_limit = std::env::var("COMMENT_LOWER_LIMIT")
        .unwrap_or_else(|_| "5".into())
        .parse::<u32>()
        .unwrap();

    println!("port: {}", port);
    println!("keyword: {}", keyword);
    println!("sort_type: {}", sort_type);
    println!("user_data_dir: {}", user_data_dir);
    println!("browser_path: {}", browser_path);
    println!("like_limit: {}", like_limit);
    println!("comment_limit: {}", comment_limit);

    let output_path = std::env::var("OUTPUT_XLSX").unwrap_or_else(|_| "xhs_notes.xlsx".into());

    let mut config = BrowserConfig::new()
        .set_local_port(port)
        .user_data_dir(&user_data_dir);

    if !browser_path.is_empty() {
        config = config.chrome_path(browser_path);
    }

    config = config.headless(false);

    let browser = Browser::connect_or_launch(config).unwrap();

    let binding = browser.tabs()?;
    let page = binding.first().unwrap();

    page.goto(&format!(
        "https://www.xiaohongshu.com/search_result?keyword={}",
        keyword
    ))?;
    random_random_sleep(1000, 2000);
    let show_filter_js = r#"
    const el = document.querySelector('.filter');
    el.dispatchEvent(new MouseEvent('mouseenter', { bubbles: true }));
    el.dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
    "#;
    page.run_js(show_filter_js)?;
    random_random_sleep(800, 1200);

    // 处理筛选条件
    let filters_elements = page.elements(".filters")?;
    for filters_element in filters_elements {
        let text = filters_element.text()?;
        if text.contains("排序依据") {
            let tags = filters_element.elements(".tags")?;
            for tag in tags {
                let tag_text = tag.text()?;
                if tag_text.contains(&sort_type) {
                    tag.click()?;
                    random_random_sleep(500, 700);
                }
            }
        }

        if text.contains("笔记类型") {
            let tags = filters_element.elements(".tags")?;
            for tag in tags {
                let tag_text = tag.text()?;
                if tag_text.contains("图文") {
                    tag.click()?;
                    random_random_sleep(500, 700);
                }
            }
        }
    }

    random_random_sleep(1000, 1200);
    // 处理卡片 收集数据
    let mut note_details = Vec::new();
    let mut handled_section_keys: HashSet<String> = HashSet::new();
    let mut need_break = false;
    loop {
        let sections = page.elements(".feeds-container section")?;
        let mut continue_scroll = true;

        for section in sections {
            // 取link为key
            let link = match section.element("a") {
                Ok(Some(link)) => link,
                _ => continue,
            };
            let href = link.attr("href")?;
            if handled_section_keys.contains(&href) {
                continue;
            }
            continue_scroll = false;
            handled_section_keys.insert(href);

            // 打开笔记
            section.elements("a")?.get(1).unwrap().click()?;
            random_random_sleep(1000, 1200);
            let note_detail = handle_detail(&page)?;
            if sort_type == "最多点赞" {
                if note_detail.like_count < like_limit {
                    need_break = true;
                }
            }
            if sort_type == "最多评论" {
                if note_detail.comment_count < comment_limit {
                    need_break = true;
                }
            }
            note_details.push(note_detail);
            // 关闭笔记
            let close_circle_ele = page.element(".close-circle")?;
            if let Some(close_circle_ele) = close_circle_ele {
                close_circle_ele.click()?;
            }
            random_random_sleep(2000, 2200);
        }

        if continue_scroll {
            handle_scroll(&page)?;
            random_random_sleep(1000, 1200);
        }
        if need_break {
            break;
        }
    }

    // 导出为xlsx文件
    let mut seen_note_links: HashSet<String> = HashSet::new();
    let filtered = note_details
        .into_iter()
        .filter(|note| seen_note_links.insert(note.note_link.clone()))
        .filter(|note| note.like_count >= like_limit && note.comment_count >= comment_limit)
        .collect::<Vec<_>>();

    if let Err(err) = export_to_xlsx(&output_path, &filtered) {
        eprintln!("导出 XLSX 失败: {err}");
    } else {
        println!("已导出 {} 条记录到 {}", filtered.len(), output_path);
    }

    return Ok(());
}
