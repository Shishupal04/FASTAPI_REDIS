from pydantic import BaseModel
from typing import Optional


class ItemCreate(BaseModel):
    key: str
    value: str


class ItemUpdate(BaseModel):
    value: str


class ItemResponse(BaseModel):
    key: str
    value: str
