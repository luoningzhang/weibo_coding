"""单元测试：sort_weibo_top_posts.py"""

import json
import os
import tempfile
import unittest

from sort_weibo_top_posts import get_engagement, load_posts, sort_top_posts


def make_post(reposts=0, likes=0, comments=0, **kwargs):
    return {
        "reposts_count": reposts,
        "attitudes_count": likes,
        "comments_count": comments,
        **kwargs,
    }


class TestGetEngagement(unittest.TestCase):
    def test_basic(self):
        post = make_post(reposts=10, likes=20, comments=5)
        self.assertEqual(get_engagement(post), 35)

    def test_zero(self):
        self.assertEqual(get_engagement({}), 0)

    def test_alternate_field_names(self):
        post = {"retweet_count": 3, "like_count": 7, "comment_count": 2}
        self.assertEqual(get_engagement(post), 12)


class TestSortTopPosts(unittest.TestCase):
    def _make_posts(self, n):
        # 总互动数依次为 1, 2, ..., n
        return [make_post(reposts=i, id=str(i)) for i in range(1, n + 1)]

    def test_returns_top_n(self):
        posts = self._make_posts(300)
        result = sort_top_posts(posts, top_n=200)
        self.assertEqual(len(result), 200)

    def test_sorted_descending(self):
        posts = self._make_posts(50)
        result = sort_top_posts(posts, top_n=10)
        totals = [get_engagement(p) for p in result]
        self.assertEqual(totals, sorted(totals, reverse=True))

    def test_top_post_is_highest(self):
        posts = self._make_posts(300)
        result = sort_top_posts(posts, top_n=200)
        self.assertEqual(get_engagement(result[0]), 300)

    def test_fewer_than_200(self):
        posts = self._make_posts(50)
        result = sort_top_posts(posts, top_n=200)
        self.assertEqual(len(result), 50)


class TestLoadPosts(unittest.TestCase):
    def _write_json(self, data):
        f = tempfile.NamedTemporaryFile(
            mode="w", suffix=".json", delete=False, encoding="utf-8"
        )
        json.dump(data, f, ensure_ascii=False)
        f.close()
        return f.name

    def test_list_format(self):
        posts = [make_post(reposts=i) for i in range(5)]
        path = self._write_json(posts)
        try:
            loaded = load_posts(path)
            self.assertEqual(len(loaded), 5)
        finally:
            os.unlink(path)

    def test_dict_with_data_key(self):
        posts = [make_post(reposts=i) for i in range(3)]
        path = self._write_json({"data": posts, "total": 3})
        try:
            loaded = load_posts(path)
            self.assertEqual(len(loaded), 3)
        finally:
            os.unlink(path)

    def test_dict_with_statuses_key(self):
        posts = [make_post(reposts=i) for i in range(2)]
        path = self._write_json({"statuses": posts})
        try:
            loaded = load_posts(path)
            self.assertEqual(len(loaded), 2)
        finally:
            os.unlink(path)

    def test_invalid_format_raises(self):
        path = self._write_json({"unknown": "value"})
        try:
            with self.assertRaises(ValueError):
                load_posts(path)
        finally:
            os.unlink(path)


if __name__ == "__main__":
    unittest.main(verbosity=2)
