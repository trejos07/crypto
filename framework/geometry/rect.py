import numpy as np

class Vector2():
	def __init__(self, x, y):
		self.x = x
		self.y = y

	def __add__(self, other):
		return Vector2(self.x + other.x, self.y + other.y)
	def __sub__(self, other):
		return Vector2(self.x - other.x, self.y - other.y)
	def __mul__(self, other):
		return Vector2(self.x * other, self.y * other)
	def __truediv__(self, other):
		return Vector2(self.x / other, self.y / other)
	def __str__(self):
		return f"({self.x}, {self.y})"
	def __repr__(self):
		return f"Vector2({self.x}, {self.y})"
	def __eq__(self, other):
		return self.x == other.x and self.y == other.y
	def __ne__(self, other):
		return self.x != other.x or self.y != other.y
	def __hash__(self):
		return hash(self.x) ^ hash(self.y)
	def __lt__(self, other):
		return self.x < other.x and self.y < other.y
	def __le__(self, other):
		return self.x <= other.x and self.y <= other.y
	def __gt__(self, other):
		return self.x > other.x and self.y > other.y
	def __ge__(self, other):
		return self.x >= other.x and self.y >= other.y
	def __neg__(self):
		return Vector2(-self.x, -self.y)
	def __abs__(self):
		return Vector2(abs(self.x), abs(self.y))
	def __bool__(self):
		return self.x != 0 or self.y != 0
	def __len__(self):
		return 2
	def __getitem__(self, key):
		if key == 0:
			return self.x
		elif key == 1:
			return self.y
		else:
			raise IndexError("Index out of range")
	def __setitem__(self, key, value):
		if key == 0:
			self.x = value
		elif key == 1:
			self.y = value
		else:
			raise IndexError("Index out of range")
	def __iter__(self):
		return iter((self.x, self.y))
	def normalize(self):
		return self / self.magnitude()
	def dot(self, other):
		return self.x * other.x + self.y * other.y
	def cross(self, other):
		return self.x * other.y - self.y * other.x
	def angle(self, other):
		return np.arccos(self.dot(other) / (self.magnitude() * other.magnitude()))
	def angle_deg(self, other):
		return np.degrees(self.angle(other))
	def magnitude(self):
		return np.sqrt(self.x**2 + self.y**2)
	def min(self, other):
		return Vector2(min(self.x, other.x), min(self.y, other.y))
	def max(self, other):
		return Vector2(max(self.x, other.x), max(self.y, other.y))

	def as_tuple(self):
		return (self.x, self.y)

class Rect():
	def __init__(self, x, y, w, h):
		self.position = Vector2(x, y)
		self.size = Vector2(w, h)

	@property
	def min(self):
		return self.position

	@min.setter
	def min(self, value):
		self.position = value
	
	@property
	def max(self):
		return self.position + self.size
	
	@max.setter
	def max(self, value):
		self.size = value - self.position
	
	@property
	def center(self):
		return Vector2(self.position.x + int(self.size.x / 2), self.position.y + int(self.size.y / 2))

	def expand(self, x, y):
		self.position -= Vector2(x, y)
		self.size += Vector2(x, y) * 2

	def encapsulate(self, other):
		new_pos = Vector2.min(self.position, other.position)
		self.size = Vector2.max(self.max, other.max) - new_pos
		self.position = new_pos
	
	@property
	def area(self):
		return self.size.x * self.size.y

	@classmethod
	def combine_rect_list(cls, rect_list : list):
		if rect_list is None or len(rect_list) == 0:
			return None

		rect = rect_list[0]
		
		for r in rect_list[1:]:
			rect.encapsulate(r)

		return rect

	def __str__(self):
		return f"{{position: {self.position}, size: {self.size}}}"

	def __repr__(self):
		return f"Rect(*{self.position}, *{self.size})"

	def copy(self):
		return eval(repr(self))