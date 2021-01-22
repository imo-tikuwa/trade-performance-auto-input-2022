class simple_encrypter:
    """
    シンプルな暗号化/複合化を行うクラス
    以下のページのコメント欄の情報を参考にクラスとして作成
    参考：https://qiita.com/magiclib/items/fe2c4b2c4a07e039b905
    """
    def __xor_string(text, key):
        if key:
            return "".join(chr(ord(c) ^ ord(key[i % len(key)])) for i, c in enumerate(text))
        else:
            return text

    @classmethod
    def encrypt(cls, text, key):
        return cls.__xor_string(text, key).encode().hex()

    @classmethod
    def decrypt(cls, text, key):
        return cls.__xor_string(bytes.fromhex(text).decode(), key)