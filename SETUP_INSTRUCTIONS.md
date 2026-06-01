# Rocket Dashboard v4 — Инструкция по подключению живых данных

## Архитектура (упрощённая)

```
Google Sheets (данные) ──► Google Apps Script (Data Proxy) ──► GitHub Pages (дашборд)
                                                                      ▲
                                                              Make.com (AI анализ, опционально)
```

**Принцип:** Apps Script читает данные напрямую из всех Google Sheets через Dashboard Config, трансформирует в JSON и отдаёт дашборду. Никаких промежуточных сервисов не нужно.

---

## Шаг 1 — Установить Google Apps Script

### 1.1 Открыть редактор
1. Откройте **Dashboard Config**: https://docs.google.com/spreadsheets/d/1EKR-czK1UvXZDIJUe5MLb70yZXyEMe1VIYTx15UOY3E
2. Меню → **Расширения** → **Apps Script**
3. Откроется редактор кода

### 1.2 Вставить код
1. Выделите весь существующий код (Ctrl+A)
2. Удалите его
3. Скопируйте содержимое файла `apps_script_proxy.js`
4. Вставьте в редактор (Ctrl+V)
5. Сохраните (Ctrl+S) — назовите проект `Rocket Dashboard Proxy`

### 1.3 Развернуть как Web App
1. Нажмите кнопку **Развернуть** (синяя кнопка справа вверху)
2. Выберите **Новое развёртывание**
3. Нажмите шестерёнку ⚙ рядом с "Тип" → выберите **Веб-приложение**
4. Настройки:
   - **Описание:** `Rocket Dashboard Data Proxy v4`
   - **Выполнять как:** Я (your@email.com)
   - **Доступ:** Все (**важно!**)
5. Нажмите **Развернуть**
6. При первом запуске — разрешите доступ (нажмите "Разрешить")
7. **Скопируйте URL** — он выглядит так:
   ```
   https://script.google.com/macros/s/AKfycb.../exec
   ```

---

## Шаг 2 — Подключить дашборд

### 2.1 Обновить index.html
В файле `index.html` найдите строки в начале скрипта:

```javascript
const C={
  DATA_PROXY:'APPS_SCRIPT_URL_PLACEHOLDER',   // ← заменить
  DEMO: true   // ← установить false
};
```

Замените на:
```javascript
const C={
  DATA_PROXY:'https://script.google.com/macros/s/ВАШ_URL/exec',
  DEMO: false
};
```

### 2.2 Залить в GitHub
```bash
git add index.html
git commit -m "Connect Apps Script Data Proxy"
git push
```

---

## Шаг 3 — Проверить работу

1. Откройте дашборд: https://epilipenko-byte.github.io/-rocket-dashboard_presentation/
2. Нажмите **↻ Обновить данные**
3. Должен появиться бейдж `● LIVE — Google Sheets`
4. Данные загрузятся из реальных Google Sheets

### Если данные не загружаются:
- Проверьте в Apps Script: **Запустить** → `testProxy` → смотрите логи
- Убедитесь что в Dashboard Config лист называется `Reports` (с большой буквы)
- Убедитесь что в Reports есть колонки: `id`, `sheet_id` (или URL), `sheet_name`, `visible`

---

## Шаг 4 — Структура Dashboard Config (лист Reports)

Лист `Reports` должен содержать колонки:

| id | name | sheet_id | sheet_name | type | visible |
|----|------|----------|------------|------|---------|
| opiu | ОПиУ | 1abc...xyz | Сводная | pnl | TRUE |
| dds_common | ДДС | 1abc...xyz | ДДС | cf | TRUE |
| balance | Баланс | 1abc...xyz | Баланс | balance | TRUE |
| traffic | Трафик | 1abc...xyz | Трафик | traffic | TRUE |
| opiu_combo | ОПиУ блоки | 1abc...xyz | комбо по блокам | pnl | TRUE |
| opiu_consol | Консолидация | 1abc...xyz | консолидация | pnl | TRUE |
| opiu_model | Модель | 1abc...xyz | Модель | model | TRUE |
| payment_cal | Платёж. календарь | 1abc...xyz | Свод | payment | TRUE |
| roadmap_ckp | ЦКП | 1abc...xyz | ЦКП | ckp | TRUE |
| roadmap_tasks | Задачи | 1abc...xyz | свод по задачам | tasks | TRUE |

**sheet_id** — это ID Google Sheets (часть URL между `/d/` и `/edit`).

---

## Шаг 5 — Добавление новых отчётов

Чтобы добавить новый отчёт:
1. Откройте Dashboard Config → лист `Reports`
2. Добавьте новую строку с данными отчёта
3. Установите `visible = TRUE`
4. Нажмите **↻ Обновить данные** в дашборде

Apps Script автоматически прочитает новый отчёт при следующем обновлении.

---

## Шаг 6 — AI Анализ (опционально, через Make.com)

Если хотите автоматический AI анализ:
1. Создайте сценарий в Make.com
2. Триггер: Webhook (принимает JSON с данными)
3. Модуль: HTTP → OpenAI/Claude API
4. Ответ: JSON с полем `analysis.text`
5. Вставьте URL webhook в `index.html` → `C.AI`

---

## Обновление данных

- **Автоматически:** Apps Script кэширует данные — каждый запрос читает свежие данные из Google Sheets
- **Вручную:** Кнопка **↻ Обновить данные** в дашборде
- **Кэш браузера:** Данные сохраняются в localStorage — при следующем открытии загружаются мгновенно

---

## Структура данных Google Sheets

Apps Script ищет строки по ключевым словам в первой колонке. Примеры:

**ОПиУ:**
- Строка "Выручка" → `revenue[]`
- Строка "Переменные расходы" → `var_costs[]`
- Строка "Операционная прибыль" → `op_profit[]`
- Строка "Чистая прибыль" → `net_profit[]`

**ДДС:**
- Строка "NCF операционный" → `ncf_op[]`
- Строка "Дивиденды" → `dividends[]`
- Строка "Остаток ДС" → `cash[]`

**Баланс:**
- Строка "Активы всего" → `assets_total[]`
- Строка "Собственный капитал" → `equity[]`
- Строка "D/E" → `de_ratio[]`

Месяцы определяются автоматически по заголовку (Янв, Фев, Мар... или Jan, Feb, Mar...).
