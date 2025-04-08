from playwright.sync_api import sync_playwright
from openpyxl import Workbook
from urllib.parse import urljoin
import time


def parse_novat():
    wb = Workbook()
    ws = wb.active
    ws.title = "НОВАТ Артисты"
    headers = ["ФИО", "Подразделение", "Должность", "URL фото", "Ссылка"]
    ws.append(headers)

    BASE_URL = "https://novat.ru"

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        page = browser.new_page()

        try:
            page.goto(f"{BASE_URL}/theatre/company", timeout=20000)
            time.sleep(3)

            department_links = page.query_selector_all("a.choose-type--white")
            departments = [{
                "name": link.inner_text().strip(),
                "url": urljoin(BASE_URL, link.get_attribute("href"))
            } for link in department_links]

            for department in departments:
                print(f"\nОбработка отдела: {department['name']}")
                page.goto(department["url"], timeout=15000)
                time.sleep(2)

                groups = page.query_selector_all("""
                    .artist-row,
                    .artist-group,
                    .team-section,
                    .staff-list
                """) or [page]

                for group in groups:
                    group_title = group.query_selector("""
                        h4.subtitle,
                        h3.subtitle,
                        .subtitle--big,
                        .subtitle--noborder,
                        .section-title
                    """)
                    group_name = group_title.inner_text().strip() if group_title else department["name"]

                    artists = group.query_selector_all("""
                        .artist,
                        .artist-group__item,
                        .staff-item,
                        .team-member,
                        [class*='artist-'],
                        [class*='person-']
                    """)

                    for artist in artists:
                        try:

                            name = None

                            name_link = artist.query_selector("a[href*='/theatre/company']")
                            if name_link:
                                name = name_link.inner_text().strip()

                            if not name:
                                name_element = artist.query_selector("""
                                    .artist__name,
                                    .artist-group__name,
                                    .name,
                                    h3, h4
                                """)
                                name = name_element.inner_text().strip() if name_element else None

                            if not name:
                                continue

                            position = ""
                            position_elements = artist.query_selector_all("""
                                .artist__position,
                                .artist-group__position,
                                .position,
                                .role,
                                .staff-position
                            """)
                            if position_elements:
                                position = " | ".join([el.inner_text().strip() for el in position_elements])

                            photo = artist.query_selector("img:not([src=''])")
                            photo_url = urljoin(BASE_URL, photo.get_attribute("src")) if photo else ""

                            profile_url = ""
                            if name_link:
                                profile_url = urljoin(BASE_URL, name_link.get_attribute("href"))
                            else:
                                profile_link = artist.query_selector("a[href*='/theatre/company']")
                                if profile_link:
                                    profile_url = urljoin(BASE_URL, profile_link.get_attribute("href"))

                            ws.append([
                                name,
                                department["name"],
                                position,
                                photo_url,
                                profile_url
                            ])
                            print(f"Добавлен: {name}")

                        except Exception as e:
                            print(f"Ошибка при обработке артиста: {str(e)[:100]}...")
                            continue

            wb.save("novat_complete_artists.xlsx")
            print("\nГотово! Все данные сохранены в novat_complete_artists.xlsx")

        except Exception as e:
            print(f"Критическая ошибка: {e}")
            page.screenshot(path="error.png")
        finally:
            browser.close()


if __name__ == "__main__":
    parse_novat()