import requests
from lxml import etree
from bs4 import BeautifulSoup
import json
from datetime import datetime, timedelta
from openpyxl import Workbook
import re


def get_courses_list():
    url = "https://www.coursera.org/sitemap~www~courses.xml"
    response = requests.get(url)
    root = etree.fromstring(response.content)
    courses_list = []
    for sitemap in root:
        url_course = sitemap.getchildren()
        courses_list.append(url_course[0].text)
    return courses_list


def get_course_info(course_slug):
    response = requests.get(course_slug)
    course_info = {}
    soup = BeautifulSoup(response.content, "lxml")
    try:
        course_info['title'] = soup.find("div", "title display-3-text").string
    except AttributeError:
        course_info['title'] = ''
    try:
        json_course_info = json.loads(
            soup.find("div", "rc-CourseGoogleSchemaMarkup").script.text)
        course_info['language'] =
            json_course_info["hasCourseInstance"][0]["inLanguage"]
        course_info['start_date'] = 
            json_course_info['hasCourseInstance'][0]["startDate"]
    except Exception:
        course_info['language'] = ''
        course_info['start_date'] = ''
    try:
         course_info['weeks'] = re.search('[0-9]+', soup.find_all(
            'div', "week-heading body-2-text")[-1].string).group()
    except AttributeError:
        course_info['weeks'] = ''
    except IndexError:
        course_info['weeks'] = ''

    try:
        course_info['average_rate'] = re.search(
            "Rating [0-9\\.\\,]+", soup.find(
                "div", "ratings-text bt3-hidden-xs").text).group()
    except IndexError:
        course_info['average_rate'] = ''
    except AttributeError:
        course_info['average_rate'] = ''
    course_info['url'] = course_slug
    return course_info


def output_courses_info_to_xlsx(filepath, course_info):
    wb_courses = Workbook()
    ws = wb_courses.active
    ws.append([
        'title',
        'language',
        'start_date'
        'weeks',
        'average_rate',
        'url'
        ])
    for course in course_info:
        ws.append([
            course['title'],
            course['language'],
            course['start_date'],
            course['weeks'],
            course['average_rate'],
            course['url']
            ])
    wb_courses.save(filepath)


if __name__ == '__main__':
    courses_list = get_courses_list()
    courses_information = []
    for course in courses_list:
        courses_information.append(get_course_info(course))
    output_courses_info_to_xlsx("./courses_list.xlsx", courses_information)
