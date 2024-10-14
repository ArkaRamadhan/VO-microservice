package models

import (
	"encoding/json"
	"time"
)

// model for memo
type Memo struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	Tanggal   *time.Time `json:"tanggal"`
	NoMemo    *string    `json:"no_memo"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

// MarshalJSON menyesuaikan serialisasi JSON untuk struct Memo
func (i *Memo) MarshalJSON() ([]byte, error) {
	type Alias Memo
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type BeritaAcara struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *BeritaAcara) MarshalJSON() ([]byte, error) {
	type Alias BeritaAcara
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type Surat struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *Surat) MarshalJSON() ([]byte, error) {
	type Alias Surat
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type Sk struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoSurat   *string    `json:"no_surat"`
	Tanggal   *time.Time `json:"tanggal"`
	Perihal   *string    `json:"perihal"`
	Pic       *string    `json:"pic"`
	CreateBy  string     `json:"create_by"`
}

func (i *Sk) MarshalJSON() ([]byte, error) {
	type Alias Sk
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

// model for perdin
type Perdin struct {
	ID        uint       `gorm:"primaryKey"`
	CreatedAt *time.Time `gorm:"autoCreateTime"`
	UpdatedAt *time.Time `gorm:"autoUpdateTime"`
	NoPerdin  *string    `json:"no_perdin"`
	Tanggal   *time.Time `json:"tanggal"`
	Hotel     *string    `json:"hotel"`
	Transport *string    `json:"transport"`
	CreateBy  string     `json:"create_by"`
}

func (i *Perdin) MarshalJSON() ([]byte, error) {
	type Alias Perdin
	if i.Tanggal == nil {
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: "", // Atau nilai default lain yang Anda inginkan
			Alias:   (*Alias)(i),
		})
	} else {
		tanggalFormatted := i.Tanggal.Format("2006-01-02")
		return json.Marshal(&struct {
			Tanggal string `json:"tanggal"`
			*Alias
		}{
			Tanggal: tanggalFormatted,
			Alias:   (*Alias)(i),
		})
	}
}

type File struct {
	ID          uint      `gorm:"primaryKey"`         // ID unik untuk file
	CreatedAt   time.Time `gorm:"autoCreateTime"`     // Timestamp saat file diunggah
	UpdatedAt   time.Time `gorm:"autoUpdateTime"`     // Timestamp untuk setiap update
	UserID      uint      `gorm:"index"`              // ID pengguna yang mengunggah file
	FilePath    string    `gorm:"not null"`           // Path lengkap di mana file disimpan
	FileName    string    `gorm:"not null"`           // Nama file asli
	ContentType string    `gorm:"not null"`           // Jenis konten file, misal 'application/pdf'
	Size        int64     `gorm:"not null"`           // Ukuran file dalam byte
}

// TableName overrides the table name used by File to `files`, if you want to specify it explicitly
func (File) TableName() string {
	return "files"
}
