// Prevents additional console window on Windows in release, DO NOT REMOVE!!
#![cfg_attr(not(debug_assertions), windows_subsystem = "windows")]

fn main() {
  tauri::Builder::default()
  .invoke_handler(tauri::generate_handler![greet, run_python_script])
    .run(tauri::generate_context!())
    .expect("error while running tauri application");
}


use std::process::{Command, Stdio};
use std::io::{BufReader, BufRead};

#[tauri::command]
fn run_python_script(file: &str) -> String{
  // 执行 Python 脚本
  let mut child = Command::new("python3")
      .args(&[file]) // 指定要执行的 Python 脚本的文件名
      .stdout(Stdio::piped()) // 捕获标准输出
      .spawn()
      .expect("Failed to execute command");

  // 从子进程的标准输出中创建一个读取器
  if let Some(stdout) = child.stdout.take() {
      let reader = BufReader::new(stdout);
      // 实时读取子进程的输出流
      for line in reader.lines() {
          println!("Python 输出: {}", line.expect("Failed to read line"));
      }
  }

  // 等待子进程结束
  let output = child.wait().expect("Failed to wait for child process");

  // 检查返回值
  if output.success() {
      println!("Python 脚本执行成功");
  } else {
      println!("Python 脚本执行失败");
  }
  format!("Hello!")
}

#[tauri::command]
fn greet(name: &str) -> String {
   format!("Hello, {}!", name)
}