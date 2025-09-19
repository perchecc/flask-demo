from flask import Flask, request
from flask_restful import Api, Resource

app = Flask(__name__)
api = Api(app)

# 模拟数据存储（重启会丢，想持久化可接 SQLite）
users = [
    {"id": 1, "name": "Alice", "email": "alice@example.com"},
    {"id": 2, "name": "Bob",   "email": "bob@example.com"},
]

class UserList(Resource):
    def get(self):
        return users          # flask-restful 会自动 jsonify

    def post(self):
        new_user = {
            "id":   len(users) + 1,
            "name": request.json.get("name"),
            "email":request.json.get("email"),
        }
        users.append(new_user)
        return new_user, 201

class User(Resource):
    def get(self, user_id):
        user = next((u for u in users if u["id"] == user_id), None)
        return (user, 200) if user else ({"message": "User not found"}, 404)

    def put(self, user_id):
        user = next((u for u in users if u["id"] == user_id), None)
        if not user:
            return {"message": "User not found"}, 404
        user["name"]  = request.json.get("name",  user["name"])
        user["email"] = request.json.get("email", user["email"])
        return user

    def delete(self, user_id):
        global users
        users = [u for u in users if u["id"] != user_id]
        return '', 204

api.add_resource(UserList, "/users")
api.add_resource(User,     "/users/<int:user_id>")

# 本地调试用；Render 用 gunicorn 启动不会走到这里
if __name__ == "__main__":
    app.run(debug=True)