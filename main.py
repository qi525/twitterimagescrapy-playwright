import os
import json
import asyncio
import httpx
from datetime import datetime

import playwright.async_api
from bs4 import BeautifulSoup
from playwright.async_api import async_playwright, expect

PROXY = "http://127.0.0.1:10808"

if not os.path.exists("images"):
    os.mkdir("images")


async def get_illustration(context, url):
    async_name = asyncio.current_task().get_name()

    with open("cookies.json", "r", encoding="utf-8") as f:
        cookies = json.load(f)

        # 如果没有指定或指定错误的 sameSite 则删除 sameSite 元素
        for cookie in cookies:
            cookie['sameSite'] = {'strict': 'Strict', 'Lax': 'lax', 'none': 'None'}.get(cookie['sameSite'])
            if cookie['sameSite'] not in ['strict', 'lax', 'none']:
                del cookie['sameSite']

        await context.add_cookies(cookies)

    page = await context.new_page()

    await page.goto(url)

    for x in range(999):  # 假设获取三次，大概 3 * 4 = 12 条
        # 等待 article 出现
        await expect(page.locator('article').nth(0)).to_be_visible(timeout=100000)  # 不建议设置太长时间的超时
        # while True:
        #     try:
        #         await expect(page.locator('article').nth(0)).to_be_visible(timeout=100000)  # 不建议设置太长时间的超时
        #     except playwright.async_api.TimeoutError as e:
        #         print(f"{async_name} ->", "获取完毕或超时。")
        #         break
        #
        #     x = 0  # 注意还要修改提示语，这里只是临时的替代方法

        # 获取 article 数量，不能直接遍历 article 因为我们采用“删除法”获取推文，直接遍历会导致元素对应错误。
        article_count = await page.locator('article').count()

        # 遍历每一条推文
        for y in range(article_count):

            print(f"{async_name} ->", "=" * 30)
            print(f"{async_name} -> 第 {x} 次循环，({y + 1}/{article_count})正在获取推文。")

            try:
                # 获取到 HTML 给 Beautiful 定位
                article = page.locator("article").first
                soup = BeautifulSoup(await article.inner_html(), "html.parser")  # lxml 可换为 html.parser

                # 简单判断是否为广告
                if 'style="text-overflow: unset;">Ad</span>' in str(soup):
                    raise ValueError("这条推文是广告。")

                # 找出相关信息
                time_element = soup.find("time")
                publish_time = time_element.get("datetime")
                publish_url = "https://x.com" + time_element.find_parent().get("href")

                # 要注意推文是否没有正文
                tweetText = soup.find("div", attrs={"data-testid": "tweetText"})
                publish_content = tweetText.get_text() if tweetText else ""

                # 还要注意是否没有图片
                tweetPhoto = soup.find("div", attrs={"data-testid": "tweetPhoto"})
                publish_images = [img.get("src") for img in tweetPhoto.find_all("img")] if tweetPhoto else []
                author = soup.find("div", attrs={"data-testid": "User-Name"}).find_all('span')[0].get_text()
                author += soup.find("div", attrs={"data-testid": "User-Name"}).find_all('span')[-2].get_text()

                # 输出信息
                print(f"{async_name} ->",
                      datetime.strptime(publish_time, "%Y-%m-%dT%H:%M:%S.%fZ").strftime(
                          "%Y年%m月%d日 %H:%M:%S"))
                print(f"{async_name} ->", "发布者：",
                      author)
                print(f"{async_name} ->", "推文地址：", publish_url)

                print(f"{async_name} ->", "推文内容：", publish_content)
                print(f"{async_name} ->", "推文图片：", publish_images)

                for image in publish_images:  # https://pbs.twimg.com/media/GQSc1-gWQAAutlr?format=png&name=small
                    # 编辑图片链接获取原图
                    image = image[:image.rfind("?")]  # https://pbs.twimg.com/media/GQSc1-gWQAAutlr
                    image_name = image[image.rfind("/") + 1:]  # GQSc1-gWQAAutlr
                    image_url = image + "?format=jpg&name=orig"  # 获取原图地址

                    async with httpx.AsyncClient(proxies={'http://': PROXY, 'https://': PROXY}) as session:
                        response = await session.get(image_url)
                        with open("images/" + image_name + ".jpg", "wb") as f:
                            f.write(response.content)


            except ValueError as e:
                print(f"{async_name} -> 第 {x} 次循环，({y + 1}/{article_count})推文获取出错：", str(e))

            # 删除推文，以找下一条
            await article.locator("xpath=../../..").first.evaluate("(element) => element.remove()")

            # 等待一会儿
            await asyncio.sleep(1)

        # await page.pause()


async def main():
    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=False)
        context = await browser.new_context()

        await asyncio.gather(
           get_illustration(context, 'https://x.com/JIN_HONG_18'),
           # get_illustration(context, 'https://x.com/mafumuffin'),
           # get_illustration(context, 'https://x.com/unko_koedame'),
           # get_illustration(context, 'https://x.com/maple926')
        )


if __name__ == '__main__':
    asyncio.run(main())
