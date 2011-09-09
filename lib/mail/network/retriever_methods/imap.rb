# encoding: utf-8

module Mail
  # The IMAP retriever allows to get the last, first or all emails from a IMAP server.
  # Each email retrieved (RFC2822) is given as an instance of +Message+.
  #
  # While being retrieved, emails can be yielded if a block is given.
  #
  # === Example of retrieving Emails from GMail:
  #
  #   Mail.defaults do
  #     retriever_method :imap, { :address             => "imap.googlemail.com",
  #                               :port                => 993,
  #                               :user_name           => '<username>',
  #                               :password            => '<password>',
  #                               :enable_ssl          => true }
  #   end
  #
  #   Mail.all    #=> Returns an array of all emails
  #   Mail.first  #=> Returns the first unread email
  #   Mail.last   #=> Returns the first unread email
  #
  # You can also pass options into Mail.find to locate an email in your imap mailbox
  # with the following options:
  #
  #   mailbox: name of the mailbox used for email retrieval. The default is 'INBOX'.
  #   what:    last or first emails. The default is :first.
  #   order:   order of emails returned. Possible values are :asc or :desc. Default value is :asc.
  #   count:   number of emails to retrieve. The default value is 10. A value of 1 returns an
  #            instance of Message, not an array of Message instances.
  #
  #   Mail.find(:what => :first, :count => 10, :order => :asc)
  #   #=> Returns the first 10 emails in ascending order
  #
  class IMAP < Retriever
    require 'net/imap'
    
    def initialize(values)
      self.settings = { :address              => "localhost",
        :port                 => 143,
        :user_name            => nil,
        :password             => nil,
        :authentication       => nil,
        :enable_ssl           => false }.merge!(values)
    end

    attr_accessor :settings

    # Find emails in a IMAP mailbox. Without any options, the 10 last received emails are returned.
    #
    # Possible options:
    #   mailbox: mailbox to search the email(s) in. The default is 'INBOX'.
    #   what:    last or first emails. The default is :first.
    #   order:   order of emails returned. Possible values are :asc or :desc. Default value is :asc.
    #   count:   number of emails to retrieve. The default value is 10. A value of 1 returns an
    #            instance of Message, not an array of Message instances.
    #   delete_after_find: flag for whether to delete each retreived email after find. Default
    #           is false. Use #find_and_delete if you would like this to default to true.
    #
    def find(options={}, &block)
      options = validate_options(options)

      start do |imap|
        imap.select(options[:mailbox])

        message_ids = imap.uid_search(options[:keys])
        message_ids.reverse! if options[:what].to_sym == :last
        message_ids = message_ids.first(options[:count]) if options[:count].is_a?(Integer)
        message_ids.reverse! if (options[:what].to_sym == :last && options[:order].to_sym == :asc) ||
          (options[:what].to_sym != :last && options[:order].to_sym == :desc)

        if block_given?
          message_ids.each do |message_id|
            fetchdata = imap.uid_fetch(message_id, ['RFC822'])[0]
            new_message = Mail.new(fetchdata.attr['RFC822'])
            new_message.mark_for_delete = true if options[:delete_after_find]
            if block.arity == 3
              yield new_message, imap, message_id
            else
              yield new_message
            end
            imap.uid_store(message_id, "+FLAGS", [Net::IMAP::DELETED]) if options[:delete_after_find] && new_message.is_marked_for_delete?
          end
          imap.expunge if options[:delete_after_find]
        else
          emails = []
          message_ids.each do |message_id|
            fetchdata = imap.uid_fetch(message_id, ['RFC822'])[0]
            emails << Mail.new(fetchdata.attr['RFC822'])
            imap.uid_store(message_id, "+FLAGS", [Net::IMAP::DELETED]) if options[:delete_after_find]
          end
          imap.expunge if options[:delete_after_find]
          emails.size == 1 && options[:count] == 1 ? emails.first : emails
        end
      end
    end

    # Delete all emails from a IMAP mailbox
    def delete_all(mailbox='INBOX')
      mailbox ||= 'INBOX'
      mailbox = Net::IMAP.encode_utf7(mailbox)

      start do |imap|
        imap.select(mailbox)
        imap.uid_search(['ALL']).each do |message_id|
          imap.uid_store(message_id, "+FLAGS", [Net::IMAP::DELETED])
        end
        imap.expunge
      end
    end

    # Returns the connection object of the retrievable (IMAP or POP3)
    def connection(&block)
      raise ArgumentError.new('Mail::Retrievable#connection takes a block') unless block_given?

      start do |imap|
        yield imap
      end
    end

    # return the status from mailbox (INBOX or other folder)
    # Supported attributes (attr) include:
    #   MESSAGES:: the number of messages in the mailbox.
    #   RECENT:: the number of recent messages in the mailbox.
    #   UNSEEN:: the number of unseen messages in the mailbox.
    def mail_status(mailbox='INBOX')
      start do |imap|
        imap.status(mailbox, ["MESSAGES"])["MESSAGES"]
      end
    end

    # get the mail from mailbox associated with the message sequence number
    # 
    # Message sequence numbers number messages within a mail box
    # from 1 up to the number of items in the mail box.  If new
    # message arrives during a session, it receives a sequence
    # number equal to the new size of the mail box.  If messages
    # are expunged from the mailbox, remaining messages have their
    # sequence numbers "shuffled down" to fill the gaps.
    # 
    # The +set+ parameter is a number or an array of
    # numbers or a Range object
    def get_mail(set,mailbox='INBOX')
      start do |imap|
        imap.select(mailbox)
        query_mail(set,imap)
      end
    end

    # this will return only the subject
    def get_header_info(set,mailbox='INBOX')
      start do |imap|
        imap.select(mailbox)
        imap.fetch(set, ["BODY[HEADER.FIELDS (SUBJECT)]","ENVELOPE"])
      end
    end

    # return the mail based on UID
    def get_mail_by_uid(uid, mailbox='INBOX')
      start do |imap|
        imap.select(mailbox)
        query_mail_by_uid(uid,imap)
      end
    end

    # given uid_array return all mails
    # returns the array of all emails
    def get_mails_by_uids(uid_array,mailbox='INBOX')
      mails = []
      start do |imap|
        imap.select(mailbox)
        uid_array.each do |uid|
          mails << query_mail_by_uid(uid,imap)
        end
      end
      mails
    end

    private

    def query_uid_mail(uid,imap_object)
      imap_object.search("HEADER Message-ID #{uid}")
    end

    def query_mail(set,imap_object)
      emails = []
      imap_object.fetch(set, ['RFC822']).each do |data|
        emails << Mail.new(data.attr['RFC822'])
      end
      emails
    end

    def query_mail_by_uid(uid,imap_object)
      # first get the sequence number of the message
      seq = query_uid_mail(uid,imap_object)
      #then get the email based on this sequence number
      if seq.nil?
        return
      else
        query_mail(seq,imap_object).at(0)
      end
    end
    
    # Set default options
    def validate_options(options)
      options ||= {}
      options[:mailbox] ||= 'INBOX'
      options[:count]   ||= 10
      options[:order]   ||= :asc
      options[:what]    ||= :first
      options[:keys]    ||= 'ALL'
      options[:delete_after_find] ||= false
      options[:mailbox] = Net::IMAP.encode_utf7(options[:mailbox])

      options
    end

    # Start an IMAP session and ensures that it will be closed in any case.
    def start(config=Mail::Configuration.instance, &block)
      raise ArgumentError.new("Mail::Retrievable#imap_start takes a block") unless block_given?

      imap = Net::IMAP.new(settings[:address], settings[:port], settings[:enable_ssl], nil, false)
      if settings[:authentication].nil?
        imap.login(settings[:user_name], settings[:password])
      else
        # Note that Net::IMAP#authenticate('LOGIN', ...) is not equal with Net::IMAP#login(...)!
        # (see also http://www.ensta.fr/~diam/ruby/online/ruby-doc-stdlib/libdoc/net/imap/rdoc/classes/Net/IMAP.html#M000718)
        imap.authenticate(settings[:authentication], settings[:user_name], settings[:password])
      end

      yield imap
    ensure
      if defined?(imap) && imap && !imap.disconnected?
        imap.disconnect
      end
    end

  end
end
