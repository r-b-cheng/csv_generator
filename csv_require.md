# CSV 文件格式与字段要求

本文档定义两类 CSV 文件的字段、字段顺序与格式要求，以便与系统的数据解析保持完全兼容。

- 学生日程：`student_schedule.csv`
- 教师办公时间：`professors.csv`

## 公共要求
- 编码：UTF-8（无 BOM）。
- 分隔符：英文逗号 `,`。
- 引号：字段中若包含逗号、换行或前后空格，需用英文双引号包裹（例如 `"Room, 101"`）。
- 日期时间格式：`YYYY-MM-DD HH:MM`（24 小时制，分钟建议用 00 或 30）。
- 工作日范围：`Weekday ∈ [1,7]`（1=周一，2=周二，…，7=周日）。
- 同一行的 `StartTime` 与 `EndTime` 必须为同一天，且满足 `EndTime > StartTime`。
- 建议确保 `Weekday` 与 `StartTime`/`EndTime` 的实际日期对应的星期一致（例如 `2025-01-06` 是周一，则 `Weekday` 应为 1），以避免后续冲突检测或周视图显示的偏差。

---

## 学生日程（student_schedule.csv）

- 文件路径建议：`example_data/student_schedule.csv`
- 字段顺序（必须严格一致）：
  1. `EventName`
  2. `Location`
  3. `Description`
  4. `Weekday`
  5. `StartTime`
  6. `EndTime`
  7. `IsCourse`

- 字段说明与约束：
  - `EventName`：字符串，必填。课程或事件名称，如 `Calculus`。
  - `Location`：字符串，必填。地点，如 `Building A Room 101`。
  - `Description`：字符串，可选。补充说明，如 `Calculus lecture`。
  - `Weekday`：整数，必填。取值 1–7，1=周一。
  - `StartTime`：日期时间，必填。格式 `YYYY-MM-DD HH:MM`。
  - `EndTime`：日期时间，必填。格式 `YYYY-MM-DD HH:MM`，且 `EndTime > StartTime`。
  - `IsCourse`：整数，必填。`1` 表示课程，`0` 表示个人事件（非课程）。

- 校验规则建议：
  - 必填字段不可为空；`Weekday` 合法；时间格式合法且顺序正确。
  - 同一日内事件时段不应重叠（冲突检查）。
  - 典型课程时长约 3 小时

- 示例（两节课，周一）：
  ```csv
  EventName,Location,Description,Weekday,StartTime,EndTime,IsCourse
  Calculus,Building A Room 101,Calculus lecture,1,2025-01-06 09:00,2025-01-06 12:00,1
  Linear Algebra,Building A Room 102,Linear algebra lecture,1,2025-01-06 15:30,2025-01-06 18:30,1
  ```

---

## 教师办公时间（professors.csv）

- 文件路径建议：`example_data/professors.csv`
- 字段顺序（必须严格一致）：
  1. `ProfessorName`
  2. `Email`
  3. `EventName`
  4. `Location`
  5. `Description`
  6. `Weekday`
  7. `StartTime`
  8. `EndTime`

- 字段说明与约束：
  - `ProfessorName`：字符串，必填。教师姓名，如 `Dr. Zhang`。
  - `Email`：字符串，必填。教师邮箱，如 `zhang@university.edu`。
  - `EventName`：字符串，必填。事件名称，如 `Office Hour`。
  - `Location`：字符串，必填。地点，如 `Room 301`。
  - `Description`：字符串，可选。说明，如 `Weekly office hour`。
  - `Weekday`：整数，必填。取值 1–7，1=周一。
  - `StartTime`：日期时间，必填。格式 `YYYY-MM-DD HH:MM`。
  - `EndTime`：日期时间，必填。格式 `YYYY-MM-DD HH:MM`，且 `EndTime > StartTime`。

- 校验规则建议：
  - 必填字段不可为空；`Email` 建议符合基础邮箱格式；
  - 保证同一名教师在同一日的办公时间不重叠；
  - 时间格式合法且顺序正确。

- 示例（周一办公时间）：
  ```csv
  ProfessorName,Email,EventName,Location,Description,Weekday,StartTime,EndTime
  Dr. Zhang,zhang@university.edu,Office Hour,Room 301,Weekly office hour,1,2025-01-06 14:00,2025-01-06 16:00
  ```

---

## 常见问题与处理建议
- 乱码/错码：请确保导出使用 UTF-8 编码；Windows 环境下避免使用 ANSI/GBK。
- 逗号干扰：字段包含逗号时需用双引号包裹；导出工具应自动处理。
- 星期不一致：若 `Weekday` 与日期对应的实际星期不一致，周视图显示与冲突检测可能出现偏差；请统一填写。
- 时间跨天：不支持跨天事件；如需跨天，请拆分为两条分别对应的日期记录。
- 行尾与空格：避免末尾多余空格或空行，以免解析异常。

---

## 模板生成建议（供工具使用）
- 课程模板（周一到周五，每天两节课）：
  - 上午：`09:00–12:00`
  - 下午：优先 `15:30–18:30`（或可选 `12:30–15:30`）
- 办公时间模板：
  - 可按教师配置每周固定的 2 小时或 3 小时段；例如周一 `14:00–16:00`。

---

## 兼容性说明
- 文档中的字段顺序与大小写需与系统解析器保持一致，否则导入将失败或产生错误数据。
- 若未来字段发生变更，请同步更新本文档并保证解析模块一致更新。
