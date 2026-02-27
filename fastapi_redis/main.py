from fastapi import FastAPI, HTTPException
from redis import Redis
from redis.exceptions import RedisError
from schemas import ItemCreate, ItemUpdate, ItemResponse

app = FastAPI()

# Redis connection
try:
    redis_client = Redis(host="localhost", port=6379, decode_responses=True)
    redis_client.ping()
except RedisError:
    redis_client = None


@app.get("/")
def test_redis():
    redis_client.set("test_key", "hello")
    value = redis_client.get("test_key")
    return {"redis_value": value}


def get_redis():
    if not redis_client:
        raise HTTPException(status_code=500, detail="Redis connection failed")
    return redis_client


# CREATE
@app.post("/items/", response_model=ItemResponse)
def create_item(item: ItemCreate):
    r = get_redis()
    try:
        if r.exists(item.key):
            raise HTTPException(status_code=400, detail="Key already exists")
        r.set(item.key, item.value)
        return ItemResponse(key=item.key, value=item.value)
    except RedisError:
        raise HTTPException(status_code=500, detail="Redis operation failed")


# READ
@app.get("/items/{key}", response_model=ItemResponse)
def read_item(key: int):
    r = get_redis()
    try:
        value = r.get(key)
        if value is None:
            raise HTTPException(status_code=404, detail="Key not found")
        return ItemResponse(key=key, value=value)
    except RedisError:
        raise HTTPException(status_code=500, detail="Redis operation failed")


# UPDATE
@app.put("/items/{key}", response_model=ItemResponse)
def update_item(key: int, item: ItemUpdate):
    r = get_redis()
    try:
        if not r.exists(key):
            raise HTTPException(status_code=404, detail="Key not found")
        r.set(key, item.value)
        return ItemResponse(key=key, value=item.value)
    except RedisError:
        raise HTTPException(status_code=500, detail="Redis operation failed")


# DELETE
@app.delete("/items/{key}")
def delete_item(key: int):
    r = get_redis()
    try:
        if not r.exists(key):
            raise HTTPException(status_code=404, detail="Key not found")
        r.delete(key)
        return {"message": "Key deleted successfully"}
    except RedisError:
        raise HTTPException(status_code=500, detail="Redis operation failed")
