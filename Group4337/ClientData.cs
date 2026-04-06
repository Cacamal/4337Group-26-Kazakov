using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Group4337
{
    public class Employee
    {
        [Key]
        public int Id { get; set; }

        [MaxLength(100)]
        public string Position { get; set; } = string.Empty;  // было Role

        [MaxLength(200)]
        public string FullName { get; set; } = string.Empty;  // новое поле

        [MaxLength(100)]
        public string Log { get; set; } = string.Empty;       // было Login

        [MaxLength(200)]
        public string Password { get; set; } = string.Empty;
    }
}
