namespace MyExercise01.Models
{
    using System;
    using System.Linq;
    using System.Collections.Generic;
    using System.ComponentModel.DataAnnotations;

    [MetadataType(typeof(客戶聯絡人MetaData))]
    public partial class 客戶聯絡人 : IValidatableObject
    {
        public IEnumerable<ValidationResult> Validate(ValidationContext validationContext)
        {
            CustomerInfoEntities db = new CustomerInfoEntities();

            if (!this.是否已刪除)
            {
                var checkEmailduplicate = db.客戶聯絡人
                                        .Where(w => w.客戶Id == this.客戶Id &&
                                                w.Email == this.Email &&
                                                w.是否已刪除 == false)
                                        .AsQueryable();

                if (this.Id == default(int))
                {
                    if (checkEmailduplicate != null && checkEmailduplicate.Count() >= 1)
                    {
                        yield return new ValidationResult("Email重複", new string[] { "Email" });
                    }
                }
                else
                {
                    checkEmailduplicate = checkEmailduplicate.Where(w => w.Id != this.Id);

                    if (checkEmailduplicate != null && checkEmailduplicate.Count() >= 1)
                    {
                        yield return new ValidationResult("Email重複", new string[] { "Email" });
                    }
                }
            }
        }
    }

    public partial class 客戶聯絡人MetaData
    {
        [Required]
        public int Id { get; set; }
        [Required]
        public int 客戶Id { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [Required]
        public string 職稱 { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [Required]
        public string 姓名 { get; set; }

        [StringLength(250, ErrorMessage = "欄位長度不得大於 250 個字元")]
        [Required]
        [EmailAddress(ErrorMessage = "信箱格式錯誤")]
        public string Email { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        [手機格式驗證]
        public string 手機 { get; set; }

        [StringLength(50, ErrorMessage = "欄位長度不得大於 50 個字元")]
        public string 電話 { get; set; }
        [Required]
        public bool 是否已刪除 { get; set; }

        public virtual 客戶資料 客戶資料 { get; set; }
    }
}
